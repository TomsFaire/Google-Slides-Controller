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

// ── pixel-format negotiation ──────────────────────────────────────────────────

test('choosePixelFormat prefers BGRA (no conversion) and falls back to YUV', () => {
  const { choosePixelFormat } = require('../src/decklink-modes');
  assert.equal(choosePixelFormat(['yuv']), 'yuv');
  assert.equal(choosePixelFormat(['yuv', 'bgra']), 'bgra');
  assert.equal(choosePixelFormat(['bgra']), 'bgra');
});

test('choosePixelFormat names the mode and what the device does support when nothing fits', () => {
  const { choosePixelFormat } = require('../src/decklink-modes');
  assert.throws(
    () => choosePixelFormat([], { mode: '1080p60', deviceModes: ['1080p30', '720p50'] }),
    /1080p60.*1080p30, 720p50/
  );
  assert.throws(() => choosePixelFormat([]), /not supported/);
});

test('filterOutputDevices drops capture-only devices but keeps their original index', () => {
  const { filterOutputDevices } = require('../src/decklink-modes');
  const out = filterOutputDevices([
    { index: 0, name: 'Recorder', canOutput: false },
    { index: 1, name: 'Monitor', canOutput: true },
    { index: 2, name: 'Unknown' },
  ]);
  assert.deepEqual(out.map((d) => d.index), [1, 2]);
});

function negotiationBackend(accepts) {
  const tried = [];
  return {
    tried,
    openOutput: async (opts) => {
      tried.push(opts.pixelFormat);
      if (!accepts.includes(opts.pixelFormat)) throw new Error('Requested display mode is not supported');
      return { displayFrame: async () => ({}), close: async () => {} };
    },
  };
}
const OPEN = { deviceIndex: 0, mode: 'bmdModeHD1080p6000', width: 4, height: 2 };

test('openOutputNegotiated tries BGRA first and falls back to YUV when the mode is not supported', async () => {
  const { openOutputNegotiated } = require('../src/decklink-backend');
  const b = negotiationBackend(['yuv']);
  const { pixelFormat } = await openOutputNegotiated(b, OPEN);
  assert.equal(pixelFormat, 'yuv');
  assert.deepEqual(b.tried, ['bgra', 'yuv']);

  const direct = negotiationBackend(['bgra', 'yuv']);
  assert.equal((await openOutputNegotiated(direct, OPEN)).pixelFormat, 'bgra');
  assert.deepEqual(direct.tried, ['bgra']);
});

test('openOutputNegotiated does not retry on errors that are not about the mode', async () => {
  const { openOutputNegotiated } = require('../src/decklink-backend');
  const tried = [];
  const b = { openOutput: async (o) => { tried.push(o.pixelFormat); throw new Error('Unable to access the hardware or input stream is currently active'); } };
  await assert.rejects(openOutputNegotiated(b, OPEN), /Unable to access the hardware/);
  assert.deepEqual(tried, ['bgra'], 'device-busy must not be masked by a second attempt');
});

test('openOutputNegotiated reports every format it tried when none work', async () => {
  const { openOutputNegotiated } = require('../src/decklink-backend');
  const b = negotiationBackend([]);
  await assert.rejects(openOutputNegotiated(b, OPEN), /not supported.*bgra, yuv|bgra, yuv.*not supported/s);
});

test('worker start negotiates the pixel format and reports the one it chose', async () => {
  const { createHandler } = require('../src/decklink-worker');
  const { replies, send } = collect();
  const b = negotiationBackend(['yuv']);
  await createHandler(b, send)({ id: 1, cmd: 'start', playbackId: 'p', ...OPEN });
  assert.deepEqual(replies[0], { id: 1, cmd: 'ok', pixelFormat: 'yuv' });
});

test('worker start honours an explicit pixelFormats list (used to force a format for colour A/B tests)', async () => {
  const { createHandler } = require('../src/decklink-worker');
  const { replies, send } = collect();
  const b = negotiationBackend(['bgra', 'yuv']);
  await createHandler(b, send)({ id: 1, cmd: 'start', playbackId: 'p', pixelFormats: ['yuv'], ...OPEN });
  assert.deepEqual(b.tried, ['yuv']);
  assert.equal(replies[0].pixelFormat, 'yuv');
});

test('WorkerProvider sends the mode\'s forced pixelFormats and remembers the negotiated format', async () => {
  const { WorkerProvider, DISPLAY_MODES } = require('../src/decklink-output');
  const sent = [];
  const worker = Object.assign(new EventEmitter(), { stderr: null, kill() {}, send(msg) {
    sent.push(msg);
    setImmediate(() => worker.emit('message', { id: msg.id, cmd: 'ok', pixelFormat: 'yuv' }));
  } });
  const mode = DISPLAY_MODES['1080p60'];
  mode.pixelFormats = ['yuv'];
  try {
    const provider = new WorkerProvider({ fork: () => worker });
    await provider.start(0, '1080p60');
    assert.deepEqual(sent[0].pixelFormats, ['yuv']);
    assert.equal(provider.pixelFormat, 'yuv');
  } finally {
    delete mode.pixelFormats;
  }
});

// ── probing the modes a device supports ───────────────────────────────────────

// Shaped like macadam's getDeviceInfo() outputDisplayModes: frameRate is [duration, scale].
const DEVICE_MODES = [
  { name: 'HD 1080p 59.94', width: 1920, height: 1080, frameRate: [1001, 60000], videoModes: ['8-bit YUV'] },
  { name: 'HD 1080p 60',    width: 1920, height: 1080, frameRate: [1000, 60000], videoModes: ['8-bit YUV', '10-bit YUV'] },
  { name: 'HD 1080p 29.97', width: 1920, height: 1080, frameRate: [1001, 30000], videoModes: ['8-bit YUV', '8-bit BGRA'] },
  { name: 'HD 1080i 59.94', width: 1920, height: 1080, frameRate: [1001, 30000], videoModes: ['8-bit BGRA'] },
  { name: 'HD 720p 60',     width: 1280, height: 720,  frameRate: [1000, 60000], videoModes: ['8-bit YUV'] },
  { name: '4K 2160p 30',    width: 3840, height: 2160, frameRate: [1000, 30000], videoModes: ['8-bit YUV'] },
  { name: 'HD 1080p 25',    width: 1920, height: 1080, frameRate: [1000, 25000], videoModes: ['10-bit YUV'] },
];

test('matchDeviceModes keeps the modes the pipeline can drive and reports their usable pixel formats', () => {
  const { matchDeviceModes } = require('../src/decklink-modes');
  const modes = matchDeviceModes(DEVICE_MODES);
  const byKey = Object.fromEntries(modes.map((m) => [m.key, m]));

  assert.deepEqual(Object.keys(byKey).sort(), ['1080i5994', '1080p2997', '1080p5994', '1080p60']);
  assert.deepEqual(byKey['1080p5994'].pixelFormats, ['yuv']);
  assert.deepEqual(byKey['1080p60'].pixelFormats, ['yuv'], '10-bit is not something the pipeline sends');
  assert.deepEqual(byKey['1080p2997'].pixelFormats, ['bgra', 'yuv']);
  assert.equal(byKey['1080p5994'].width, 1920);
  assert.match(byKey['1080p5994'].label, /1080p.*59\.94/);
});

test('matchDeviceModes tells 1080i59.94 from 1080p29.97 even though both report 29.97 fps', () => {
  const { matchDeviceModes } = require('../src/decklink-modes');
  const both = matchDeviceModes(DEVICE_MODES);
  assert.equal(both.find((m) => m.key === '1080i5994').pixelFormats.join(), 'bgra');
  assert.equal(both.find((m) => m.key === '1080p2997').pixelFormats.join(), 'bgra,yuv');

  const onlyInterlaced = matchDeviceModes([DEVICE_MODES[3]]);
  assert.deepEqual(onlyInterlaced.map((m) => m.key), ['1080i5994']);
  const onlyProgressive = matchDeviceModes([DEVICE_MODES[2]]);
  assert.deepEqual(onlyProgressive.map((m) => m.key), ['1080p2997']);
});

test('matchDeviceModes drops modes with no pixel format we can send, and copes with junk', () => {
  const { matchDeviceModes } = require('../src/decklink-modes');
  assert.equal(matchDeviceModes([DEVICE_MODES[6]]).length, 0, '1080p25 only offers 10-bit YUV');
  assert.deepEqual(matchDeviceModes([]), []);
  assert.deepEqual(matchDeviceModes(undefined), []);
  assert.deepEqual(matchDeviceModes([{ name: 'x' }]), []);
});

test('macadam backend listModes reads the requested device and returns [] for an unknown index', async () => {
  const { createMacadamBackend } = require('../src/backends/macadam-backend');
  const mod = { getDeviceInfo: async () => [{ outputDisplayModes: [DEVICE_MODES[0]] }, { outputDisplayModes: [DEVICE_MODES[2]] }] };
  const backend = createMacadamBackend(() => mod);
  assert.deepEqual((await backend.listModes(1)).map((m) => m.key), ['1080p2997']);
  assert.deepEqual((await backend.listModes(0)).map((m) => m.key), ['1080p5994']);
  assert.deepEqual(await backend.listModes(7), []);
});

test('worker get_modes asks the backend for the requested device', async () => {
  const { createHandler } = require('../src/decklink-worker');
  const { replies, send } = collect();
  const asked = [];
  const backend = { listModes: async (i) => { asked.push(i); return [{ key: '1080p60' }]; } };
  await createHandler(backend, send)({ id: 4, cmd: 'get_modes', deviceIndex: 2 });
  assert.deepEqual(asked, [2]);
  assert.deepEqual(replies[0], { id: 4, cmd: 'ok', modes: [{ key: '1080p60' }] });
});

// ── choose the format from what the device reports, not by trial and error ────

test('matchDeviceModes includes the BMD mode name so a start request can be matched to it', () => {
  const { matchDeviceModes } = require('../src/decklink-modes');
  const m = matchDeviceModes(DEVICE_MODES).find((x) => x.key === '1080p60');
  assert.equal(m.bmdMode, 'bmdModeHD1080p6000');
});

test('worker start opens the device once, in the format the device reports for that mode', async () => {
  const { createHandler } = require('../src/decklink-worker');
  const { replies, send } = collect();
  const b = negotiationBackend(['bgra', 'yuv']);
  b.listModes = async () => [{ key: '1080p60', bmdMode: 'bmdModeHD1080p6000', pixelFormats: ['yuv'] }];
  await createHandler(b, send)({ id: 1, cmd: 'start', playbackId: 'p', bmdMode: 'bmdModeHD1080p6000', ...OPEN });
  assert.deepEqual(b.tried, ['yuv'], 'no rejected BGRA attempt before the real open');
  assert.equal(replies[0].pixelFormat, 'yuv');
});

test('worker start falls back to trying each format when the device mode list is unavailable', async () => {
  const { createHandler } = require('../src/decklink-worker');
  for (const listModes of [async () => { throw new Error('boom'); }, async () => [], async () => [{ bmdMode: 'other', pixelFormats: ['bgra'] }]]) {
    const { replies, send } = collect();
    const b = negotiationBackend(['yuv']);
    b.listModes = listModes;
    await createHandler(b, send)({ id: 1, cmd: 'start', playbackId: 'p', ...OPEN });
    assert.deepEqual(b.tried, ['bgra', 'yuv']);
    assert.equal(replies[0].pixelFormat, 'yuv');
  }
});
