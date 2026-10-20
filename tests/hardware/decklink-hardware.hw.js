'use strict';

/**
 * Hardware integration test for a real DeckLink device
 * (e.g. UltraStudio Express / Monitor 3G).
 *
 *   DECKLINK_HW=1 yarn test:decklink-hw
 *   DECKLINK_HW=1 DECKLINK_DEVICE=1 DECKLINK_MODE=1080p2997 yarn test:decklink-hw
 *
 * Runs the real worker + provider pipeline against fake windows that serve
 * SMPTE-style colour bars, so no Electron or Google session is needed.
 * Runs under Electron's Node (ELECTRON_RUN_AS_NODE) because macadam is
 * rebuilt for Electron, matching how the app's worker loads it.
 * Needs macadam built (`yarn install && yarn rebuild`) and Desktop Video
 * installed. Look at the device's SDI/HDMI out: bars on test 3, a black
 * flash on test 4.
 *
 * Not named *.test.js on purpose, so `yarn test` never picks it up.
 */

const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const { DecklinkOutputManager, DISPLAY_MODES } = require('../../src/decklink-output');

const ENABLED = process.env.DECKLINK_HW === '1';
const DEVICE_INDEX = Number(process.env.DECKLINK_DEVICE || 0);
const MODE = process.env.DECKLINK_MODE || '1080p60';
const HOLD_MS = Number(process.env.DECKLINK_HOLD_MS || 4000);
const skip = ENABLED ? false : 'set DECKLINK_HW=1 to run against real hardware';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// 7 vertical colour bars in BGRA at the requested mode size.
function colourBars(width, height) {
  const bars = [
    [255, 255, 255], [255, 255, 0], [0, 255, 255], [0, 255, 0],
    [255, 0, 255], [255, 0, 0], [0, 0, 255],
  ];
  const buf = Buffer.alloc(width * height * 4);
  for (let x = 0; x < width; x++) {
    const [r, g, b] = bars[Math.floor((x / width) * bars.length)];
    for (let y = 0; y < height; y++) {
      const o = (y * width + x) * 4;
      buf[o] = b; buf[o + 1] = g; buf[o + 2] = r; buf[o + 3] = 255;
    }
  }
  return buf;
}

// Minimal stand-in for a BrowserWindow, as used by OutputController._tick.
function fakeWindow(width, height) {
  const bitmap = colourBars(width, height);
  let captures = 0;
  const image = { resize: () => ({ toBitmap: () => bitmap }) };
  const win = {
    isDestroyed: () => false,
    webContents: { capturePage: async () => { captures++; return image; } },
  };
  return { win, captures: () => captures };
}

// Fail loudly if we silently fell back to ffmpeg/unavailable.
function assertMacadam() {
  const { providerType, detectionErrors } = DecklinkOutputManager.getStatus();
  assert.equal(providerType, 'macadam',
    `provider was ${providerType}: ${JSON.stringify(detectionErrors)}`);
}

const prefsFor = (enabled) => ({
  decklink: { slides: { enabled, deviceIndex: DEVICE_INDEX, displayMode: MODE }, notes: { enabled: false } },
});

after(async () => { await DecklinkOutputManager.shutdown(); });

test('display mode under test is defined', { skip }, () => {
  assert.ok(DISPLAY_MODES[MODE], `unknown DECKLINK_MODE ${MODE}`);
});

test('macadam provider is detected', { skip }, async () => {
  await DecklinkOutputManager.init(() => null, () => null, prefsFor(false));
  const status = DecklinkOutputManager.getStatus();
  assert.equal(status.providerType, 'macadam',
    `provider was ${status.providerType}: ${JSON.stringify(status.detectionErrors)}`);
});

test('device list includes the target device', { skip }, async () => {
  const devices = await DecklinkOutputManager.getDevices();
  console.log('  devices:', devices.map((d) => `${d.index}:${d.name}`).join(', ') || '(none)');
  assert.ok(devices.length > DEVICE_INDEX,
    `expected a device at index ${DEVICE_INDEX}, found ${devices.length}`);
});

test('slides output streams colour bars without errors', { skip }, async () => {
  const { width, height } = DISPLAY_MODES[MODE];
  const { win, captures } = fakeWindow(width, height);
  await DecklinkOutputManager.init(() => win, () => null, prefsFor(false));
  await DecklinkOutputManager.reconfigure(prefsFor(true));

  assertMacadam();
  let status = DecklinkOutputManager.getStatus();
  assert.equal(status.slides.error, null, `start failed: ${status.slides.error}`);
  assert.equal(status.slides.active, true);

  await sleep(HOLD_MS);
  status = DecklinkOutputManager.getStatus();
  assert.equal(status.slides.active, true, 'output stopped during hold');
  assert.equal(status.slides.error, null);

  const expected = (HOLD_MS / 1000) * DISPLAY_MODES[MODE].fps;
  assert.ok(captures() > expected * 0.5,
    `only ${captures()} frames captured in ${HOLD_MS}ms (expected ~${Math.round(expected)})`);
});

test('missing window outputs black without crashing', { skip }, async () => {
  await DecklinkOutputManager.reconfigure(prefsFor(false));
  await DecklinkOutputManager.init(() => null, () => null, prefsFor(true));
  await sleep(2000);
  assertMacadam();
  const status = DecklinkOutputManager.getStatus();
  assert.equal(status.slides.active, true);
  assert.equal(status.slides.error, null);
});

test('output can be stopped and restarted on the same device', { skip }, async () => {
  const { width, height } = DISPLAY_MODES[MODE];
  const { win } = fakeWindow(width, height);
  for (let i = 0; i < 3; i++) {
    await DecklinkOutputManager.reconfigure(prefsFor(false));
    assert.equal(DecklinkOutputManager.getStatus().slides.active, false);
    await DecklinkOutputManager.init(() => win, () => null, prefsFor(true));
    assertMacadam();
    const status = DecklinkOutputManager.getStatus();
    assert.equal(status.slides.error, null, `restart ${i} failed: ${status.slides.error}`);
    await sleep(1000);
  }
});

test('shutdown stops all outputs', { skip }, async () => {
  await DecklinkOutputManager.shutdown();
  const status = DecklinkOutputManager.getStatus();
  assert.equal(status.slides.active, false);
  assert.equal(status.notes.active, false);
});
