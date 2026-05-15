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
  const expected = ['1080p2997', '1080p25', '1080p30', '1080i5994', '720p5994', '720p50'];
  for (const key of expected) {
    assert.ok(key in DISPLAY_MODES, `DISPLAY_MODES should contain ${key}`);
    assert.ok(DISPLAY_MODES[key].width > 0, `${key} should have positive width`);
    assert.ok(DISPLAY_MODES[key].height > 0, `${key} should have positive height`);
    assert.ok(DISPLAY_MODES[key].fps > 0, `${key} should have positive fps`);
  }
});
