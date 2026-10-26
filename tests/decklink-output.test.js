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
  const expected = ['1080p5994', '1080p60', '1080p50', '1080p2997', '1080p25', '1080p30', '720p5994', '720p50'];
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
