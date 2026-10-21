const { test } = require('node:test');
const assert = require('node:assert/strict');

// Helper: attempt to require the module, returning null if unavailable
function tryRequireDecklinkOutput() {
  try {
    return require('../src/decklink-output');
  } catch (e) {
    return null;
  }
}

test('DecklinkOutputManager loads without crashing when macadam unavailable', async () => {
  const mod = tryRequireDecklinkOutput();
  assert.ok(mod !== null, 'module should be requireable (src/decklink-output.js must exist)');
  const { DecklinkOutputManager } = mod;
  assert.ok(DecklinkOutputManager, 'DecklinkOutputManager should be exported');
});

test('getStatus returns unavailable shape when not initialized', () => {
  const mod = tryRequireDecklinkOutput();
  if (!mod) return; // skip if module not yet present
  const { DecklinkOutputManager } = mod;
  const status = DecklinkOutputManager.getStatus();
  assert.ok(typeof status === 'object', 'status should be an object');
  assert.ok('providerType' in status, 'status should have providerType');
  assert.ok('slides' in status, 'status should have slides');
  assert.ok('notes' in status, 'status should have notes');
  assert.equal(typeof status.slides.active, 'boolean');
  assert.equal(typeof status.notes.active, 'boolean');
});

test('init with both outputs disabled completes without error', async () => {
  const mod = tryRequireDecklinkOutput();
  if (!mod) return;
  const { DecklinkOutputManager } = mod;
  const prefs = {
    decklink: {
      slides: { enabled: false, deviceIndex: 0, displayMode: '1080p2997' },
      notes:  { enabled: false, deviceIndex: 1, displayMode: '1080p2997' }
    }
  };
  await assert.doesNotReject(
    DecklinkOutputManager.init(() => null, () => null, prefs)
  );
});

test('init with no decklink prefs completes without error', async () => {
  const mod = tryRequireDecklinkOutput();
  if (!mod) return;
  const { DecklinkOutputManager } = mod;
  await assert.doesNotReject(
    DecklinkOutputManager.init(() => null, () => null, {})
  );
});

test('shutdown completes without error when nothing started', async () => {
  const mod = tryRequireDecklinkOutput();
  if (!mod) return;
  const { DecklinkOutputManager } = mod;
  await assert.doesNotReject(DecklinkOutputManager.shutdown());
});

test('getDevices returns an array', async () => {
  const mod = tryRequireDecklinkOutput();
  if (!mod) return;
  const { DecklinkOutputManager } = mod;
  const devices = await DecklinkOutputManager.getDevices();
  assert.ok(Array.isArray(devices), 'getDevices should return an array');
});

test('reconfigure with valid prefs does not throw', async () => {
  const mod = tryRequireDecklinkOutput();
  if (!mod) return;
  const { DecklinkOutputManager } = mod;
  const prefs = {
    decklink: {
      slides: { enabled: false, deviceIndex: 0, displayMode: '1080p25' },
      notes:  { enabled: false, deviceIndex: 1, displayMode: '1080p25' }
    }
  };
  await assert.doesNotReject(DecklinkOutputManager.reconfigure(prefs));
});

test('DISPLAY_MODES contains expected keys', () => {
  let DISPLAY_MODES;
  try {
    DISPLAY_MODES = require('../src/decklink-output').DISPLAY_MODES;
  } catch (e) { return; }
  if (!DISPLAY_MODES) return;
  const expected = ['1080p5994', '1080p60', '1080p50', '1080p2997', '1080p25', '1080p30', '1080i5994', '720p5994', '720p50'];
  for (const key of expected) {
    assert.ok(key in DISPLAY_MODES, `DISPLAY_MODES should contain ${key}`);
    assert.ok(DISPLAY_MODES[key].width > 0, `${key} should have positive width`);
    assert.ok(DISPLAY_MODES[key].height > 0, `${key} should have positive height`);
    assert.ok(DISPLAY_MODES[key].fps > 0, `${key} should have positive fps`);
  }
});

test('OutputController never runs overlapping captures (no backlog behind a slow capturePage)', async () => {
  const mod = tryRequireDecklinkOutput();
  assert.ok(mod && mod.OutputController, 'OutputController should be exported');

  let inFlight = 0;
  let maxInFlight = 0;
  let captures = 0;
  const image = { getSize: () => ({ width: 1920, height: 1080 }), resize: () => image, toBitmap: () => Buffer.alloc(4) };
  const win = {
    isDestroyed: () => false,
    webContents: {
      capturePage: async () => {
        captures++;
        inFlight++;
        maxInFlight = Math.max(maxInFlight, inFlight);
        await new Promise((r) => setTimeout(r, 50)); // slower than a 59.94fps tick (~17ms)
        inFlight--;
        return image;
      },
    },
  };
  const pushed = [];
  const provider = { pushFrame: (b) => { pushed.push(b); return true; } };

  const c = new mod.OutputController();
  c.start(() => win, provider, '1080p5994', 'test');
  await new Promise((r) => setTimeout(r, 500));
  c.stop();
  await new Promise((r) => setTimeout(r, 100)); // let the last capture settle

  assert.equal(maxInFlight, 1, 'captures must not overlap');
  assert.ok(captures <= 11, `expected ~10 captures in 500ms at 50ms each, got ${captures}`);
  assert.ok(pushed.length >= 1, 'frames should still be pushed');
});

test('OutputController counts frames the device could not accept', async () => {
  const mod = tryRequireDecklinkOutput();
  assert.ok(mod && mod.OutputController);
  const image = { getSize: () => ({ width: 1920, height: 1080 }), resize: () => image, toBitmap: () => Buffer.alloc(4) };
  const win = { isDestroyed: () => false, webContents: { capturePage: async () => image } };
  const c = new mod.OutputController();
  c.start(() => win, { pushFrame: () => false }, '1080p5994', 'test');
  await new Promise((r) => setTimeout(r, 120));
  c.stop();
  assert.ok(c._stats.droppedPushes >= 1);
});

// ── FFmpeg fallback ───────────────────────────────────────────────────────────

const { EventEmitter } = require('node:events');

function fakeFfmpegProc() {
  const proc = new EventEmitter();
  proc.stdin  = Object.assign(new EventEmitter(), { writable: true, writes: [], write(b, cb) { this.writes.push(b); if (cb) cb(); } });
  proc.stderr = new EventEmitter();
  proc.killed = false;
  proc.kill   = () => { proc.killed = true; };
  return proc;
}

test('ffmpegListsDecklink only matches builds that list a decklink muxer', () => {
  const { ffmpegListsDecklink } = require('../src/decklink-output');
  const homebrew = 'Devices:\n D. = Demuxing supported\n .E = Muxing supported\n --\n D  avfoundation   AVFoundation input device\n D  lavfi          Libavfilter virtual input device\n';
  const withDeck = homebrew + ' DE decklink        Blackmagic DeckLink output\n';
  const outOnly  = homebrew + '  E decklink        Blackmagic DeckLink output\n';
  assert.equal(ffmpegListsDecklink(homebrew), false);
  assert.equal(ffmpegListsDecklink(withDeck), true);
  assert.equal(ffmpegListsDecklink(outOnly), true);
  assert.equal(ffmpegListsDecklink(''), false);
  assert.equal(ffmpegListsDecklink(undefined), false);
});

test('detectFfmpeg reports why ffmpeg is unusable instead of assuming it works', () => {
  const { detectFfmpeg } = require('../src/decklink-output');
  const missing = () => { const e = new Error('spawn ffmpeg ENOENT'); e.code = 'ENOENT'; throw e; };
  assert.deepEqual(detectFfmpeg(missing), { available: false, error: 'ffmpeg not found' });

  const noDeck = detectFfmpeg(() => ' D  lavfi  Libavfilter\n');
  assert.equal(noDeck.available, false);
  assert.match(noDeck.error, /without DeckLink support/);

  const crashed = detectFfmpeg(() => { throw new Error('boom'); });
  assert.deepEqual(crashed, { available: false, error: 'boom' });

  assert.deepEqual(detectFfmpeg(() => ' DE decklink  Blackmagic\n'), { available: true, error: null });
});

test('parseFfmpegSinks extracts device names and ignores the header', () => {
  const { parseFfmpegSinks } = require('../src/decklink-output');
  const out = 'Auto-detected sinks for decklink:\n  UltraStudio Express Monitor 3G [UltraStudio Express Monitor 3G]\n* DeckLink Mini Monitor [DeckLink Mini Monitor]\n';
  assert.deepEqual(parseFfmpegSinks(out), ['UltraStudio Express Monitor 3G', 'DeckLink Mini Monitor']);
  assert.deepEqual(parseFfmpegSinks('Auto-detected sinks for decklink:\n'), []);
  assert.deepEqual(parseFfmpegSinks(''), []);
});

test('ffmpeg frame rates use exact rationals for NTSC rates', () => {
  const { ffmpegFrameRate } = require('../src/decklink-output');
  assert.equal(ffmpegFrameRate(59.94), '60000/1001');
  assert.equal(ffmpegFrameRate(29.97), '30000/1001');
  assert.equal(ffmpegFrameRate(60), '60');
  assert.equal(ffmpegFrameRate(25), '25');
});

test('buildFfmpegArgs targets the named device at the mode size and rate', () => {
  const { buildFfmpegArgs, DISPLAY_MODES } = require('../src/decklink-output');
  const args = buildFfmpegArgs(DISPLAY_MODES['1080p5994'], 'My Device');
  assert.equal(args[args.indexOf('-video_size') + 1], '1920x1080');
  assert.equal(args[args.indexOf('-framerate') + 1], '60000/1001');
  assert.equal(args[args.length - 1], 'My Device');
});

test('FfmpegProvider.start rejects with ffmpeg\'s own error when it exits straight away', async () => {
  const { FfmpegProvider } = require('../src/decklink-output');
  const proc = fakeFfmpegProc();
  const p = new FfmpegProvider({ spawn: () => proc, deviceName: () => 'X', graceMs: 1000 });
  const started = p.start(0, '1080p5994');
  proc.stderr.emit('data', Buffer.from('[decklink @ 0x1] Could not open output\nUnsupported frame rate\n'));
  proc.emit('exit', 1, null);
  await assert.rejects(started, /exited \(code=1.*Unsupported frame rate/);
  assert.equal(p.pushFrame(Buffer.alloc(4)), false, 'dead process must not accept frames');
});

test('FfmpegProvider.start rejects when ffmpeg cannot be spawned', async () => {
  const { FfmpegProvider } = require('../src/decklink-output');
  const proc = fakeFfmpegProc();
  const p = new FfmpegProvider({ spawn: () => proc, deviceName: () => 'X', graceMs: 1000 });
  const started = p.start(0, '1080p5994');
  proc.emit('error', new Error('spawn ffmpeg ENOENT'));
  await assert.rejects(started, /failed to start: spawn ffmpeg ENOENT/);
});

test('FfmpegProvider survives ffmpeg dying mid-stream (EPIPE) and reports busy as a drop', async () => {
  const { FfmpegProvider } = require('../src/decklink-output');
  const proc = fakeFfmpegProc();
  const p = new FfmpegProvider({ spawn: () => proc, deviceName: () => 'X', graceMs: 5 });
  await p.start(0, '1080p5994');

  assert.equal(p.pushFrame(Buffer.alloc(4)), true);
  proc.stdin.write = function (b) { this.writes.push(b); };   // write never completes: busy
  assert.equal(p.pushFrame(Buffer.alloc(4)), true);
  assert.equal(p.pushFrame(Buffer.alloc(4)), false, 'busy push must report false so drops are counted');

  assert.doesNotThrow(() => proc.stdin.emit('error', new Error('write EPIPE')));
  proc.emit('exit', 1, null);
  assert.equal(p.pushFrame(Buffer.alloc(4)), false, 'no writes after ffmpeg has exited');
});
