const { test } = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');

// ── createBackend ─────────────────────────────────────────────────────────────

test('createBackend returns a macadam backend by default and rejects unknown names', () => {
  const { createBackend } = require('../src/decklink-backend');
  const b = createBackend('macadam');
  for (const fn of ['probe', 'listDevices', 'openOutput']) {
    assert.equal(typeof b[fn], 'function', `backend.${fn} should be a function`);
  }
  assert.throws(() => createBackend('nope'), /unknown DeckLink backend: nope/);
});

test('createBackend reads DECKLINK_BACKEND from the environment when no name is given', () => {
  const { createBackend } = require('../src/decklink-backend');
  const prev = process.env.DECKLINK_BACKEND;
  try {
    process.env.DECKLINK_BACKEND = 'nope';
    assert.throws(() => createBackend(), /unknown DeckLink backend: nope/);
    delete process.env.DECKLINK_BACKEND;
    assert.doesNotThrow(() => createBackend());
  } finally {
    if (prev === undefined) delete process.env.DECKLINK_BACKEND;
    else process.env.DECKLINK_BACKEND = prev;
  }
});

// ── macadam backend against a fake macadam module ─────────────────────────────

function fakeMacadam(overrides = {}) {
  const calls = { playback: [], frames: [], stopped: 0 };
  const mod = {
    bmdModeHD1080p2997: 1, bmdModeHD1080p5994: 2,
    bmdFormat8BitBGRA: 'BGRA', bmdFormat8BitYUV: 'YUV',
    deckLinkVersion: () => 'DeckLinkAPI version: 16.4.0',
    getDeviceInfo: async () => [{ displayName: 'UltraStudio' }, { modelName: 'Model Only' }, {}],
    playback: async (opts) => {
      calls.playback.push(opts);
      return {
        displayFrame: async (buf) => { calls.frames.push(buf); },
        stop: async () => { calls.stopped++; },
      };
    },
    ...overrides,
  };
  return { mod, calls };
}

test('macadam backend probe reports the driver version and fails when constants are missing', async () => {
  const { createMacadamBackend } = require('../src/backends/macadam-backend');
  const { mod } = fakeMacadam();
  assert.deepEqual(await createMacadamBackend(() => mod).probe(), { driverVersion: 'DeckLink API 16.4.0' });

  const noDriver = fakeMacadam({ deckLinkVersion: () => { throw new Error('Unable to load DeckLinkAPI.'); } });
  assert.deepEqual(await createMacadamBackend(() => noDriver.mod).probe(), { driverVersion: null });

  const stale = fakeMacadam({ bmdModeHD1080p2997: undefined });
  await assert.rejects(createMacadamBackend(() => stale.mod).probe(), /rebuild required/);
});

test('macadam backend lists devices with a name fallback', async () => {
  const { createMacadamBackend } = require('../src/backends/macadam-backend');
  const { mod } = fakeMacadam();
  const devices = await createMacadamBackend(() => mod).listDevices();
  assert.deepEqual(devices.map((d) => d.name), ['UltraStudio', 'Model Only', 'DeckLink 2']);
  assert.deepEqual(devices.map((d) => d.index), [0, 1, 2]);
});

test('macadam backend sends BGRA frames through unchanged and converts for YUV outputs', async () => {
  const { createMacadamBackend } = require('../src/backends/macadam-backend');
  const { mod, calls } = fakeMacadam();
  const backend = createMacadamBackend(() => mod);

  const bgraOut = await backend.openOutput({ deviceIndex: 0, mode: 'bmdModeHD1080p2997', width: 4, height: 2, pixelFormat: 'bgra' });
  assert.equal(calls.playback[0].pixelFormat, 'BGRA');
  assert.equal(calls.playback[0].displayMode, 1);
  const frame = Buffer.alloc(4 * 2 * 4, 0x80);
  const t = await bgraOut.displayFrame(frame);
  assert.equal(calls.frames[0], frame, 'BGRA frame is passed through as-is');
  assert.equal(typeof t.convertMs, 'number');
  assert.equal(typeof t.displayMs, 'number');

  const yuvOut = await backend.openOutput({ deviceIndex: 0, mode: 'bmdModeHD1080p5994', width: 4, height: 2, pixelFormat: 'yuv' });
  assert.equal(calls.playback[1].pixelFormat, 'YUV');
  await yuvOut.displayFrame(frame);
  assert.equal(calls.frames[1].length, 4 * 2 * 2, 'UYVY is 2 bytes per pixel');

  await yuvOut.close();
  assert.equal(calls.stopped, 1);
});

// ── worker message handler ────────────────────────────────────────────────────

function fakeBackend({ displayFrame, openOutput } = {}) {
  const output = {
    displayFrame: displayFrame || (async () => ({ convertMs: 1, displayMs: 2 })),
    close: async () => {},
  };
  return {
    probe: async () => ({ driverVersion: 'DeckLink API 16.4.0' }),
    listDevices: async () => [{ index: 0, name: 'X', canOutput: true }],
    openOutput: openOutput || (async () => output),
  };
}

function collect() {
  const replies = [];
  return { replies, send: (id, fields) => replies.push({ id, ...fields }) };
}

const START = { id: 1, cmd: 'start', playbackId: 'p1', deviceIndex: 0, bmdMode: 'bmdModeHD1080p2997', pixelFormat: 'bgra', width: 4, height: 2 };

test('worker handler replies with the backend timings for a frame', async () => {
  const { createHandler } = require('../src/decklink-worker');
  const { replies, send } = collect();
  const handle = createHandler(fakeBackend(), send);
  await handle(START);
  await handle({ id: 2, cmd: 'frame', playbackId: 'p1', data: new ArrayBuffer(32) });
  assert.deepEqual(replies[1], { id: 2, cmd: 'ok', timing: { convertMs: 1, displayMs: 2 } });
});

test('worker handler surfaces the backend error when the output cannot be opened', async () => {
  const { createHandler } = require('../src/decklink-worker');
  const { send } = collect();
  const handle = createHandler(fakeBackend({
    openOutput: async () => { throw new Error('Requested display mode is not supported'); },
  }), send);
  await assert.rejects(handle(START), /Requested display mode is not supported/);
});

test('worker handler skips a frame that arrives while the previous one is still being displayed', async () => {
  const { createHandler } = require('../src/decklink-worker');
  const { replies, send } = collect();
  let release;
  const gate = new Promise((r) => { release = r; });
  const handle = createHandler(fakeBackend({ displayFrame: async () => { await gate; return { convertMs: 0, displayMs: 0 }; } }), send);
  await handle(START);
  const first = handle({ id: 2, cmd: 'frame', playbackId: 'p1', data: new ArrayBuffer(32) });
  await handle({ id: 3, cmd: 'frame', playbackId: 'p1', data: new ArrayBuffer(32) });
  assert.deepEqual(replies.find((r) => r.id === 3), { id: 3, cmd: 'skip' });
  release();
  await first;
  assert.equal(replies.find((r) => r.id === 2).cmd, 'ok');
});

test('worker handler reports the driver version from probe', async () => {
  const { createHandler } = require('../src/decklink-worker');
  const { replies, send } = collect();
  await createHandler(fakeBackend(), send)({ id: 1, cmd: 'probe' });
  assert.deepEqual(replies[0], { id: 1, cmd: 'ok', driverVersion: 'DeckLink API 16.4.0' });
});

// ── provider: dead worker ─────────────────────────────────────────────────────

test('WorkerProvider stops accepting frames once its worker has exited', async () => {
  const { WorkerProvider } = require('../src/decklink-output');
  const worker = Object.assign(new EventEmitter(), { stderr: null, kill() {}, send(msg) {
    // answer the start request like a live worker would
    setImmediate(() => worker.emit('message', { id: msg.id, cmd: 'ok' }));
  } });
  const provider = new WorkerProvider({ fork: () => worker });
  await provider.start(0, '1080p2997');
  assert.equal(provider.pushFrame(Buffer.alloc(8)), true);
  worker.emit('exit', 1, null);
  assert.equal(provider.pushFrame(Buffer.alloc(8)), false, 'no frames to a dead worker');
});
