'use strict';

/**
 * Tries every DISPLAY_MODES entry on a real DeckLink device and reports which
 * ones the hardware accepts. Each supported mode streams colour bars for
 * HOLD_MS so you can see it on the output.
 *
 *   yarn decklink:probe-modes
 *   DECKLINK_ONLY=1080p25 DECKLINK_HOLD_MS=8000 yarn decklink:probe-modes
 *   DECKLINK_DEVICE=1 DECKLINK_HOLD_MS=3000 yarn decklink:probe-modes
 *
 * Every mode runs in its own fresh process so one attempt cannot leave the
 * device in a state that affects the next. DECKLINK_ONLY=<mode> runs just
 * that mode in this process.
 *
 * Runs under ELECTRON_RUN_AS_NODE (see package.json) because macadam is
 * rebuilt for Electron.
 */

const { spawnSync } = require('child_process');
const { DecklinkOutputManager, DISPLAY_MODES } = require('../src/decklink-output');

const DEVICE_INDEX = Number(process.env.DECKLINK_DEVICE || 0);
const HOLD_MS = Number(process.env.DECKLINK_HOLD_MS || 3000);
const GAP_MS = 2500;
const ONLY = process.env.DECKLINK_ONLY;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function colourBars(width, height) {
  const bars = [[255, 255, 255], [255, 255, 0], [0, 255, 255], [0, 255, 0],
    [255, 0, 255], [255, 0, 0], [0, 0, 255]];
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

function fakeWindow(width, height) {
  const bitmap = colourBars(width, height);
  const image = { resize: () => ({ toBitmap: () => bitmap }) };
  return { isDestroyed: () => false, webContents: { capturePage: async () => image } };
}

const prefsFor = (displayMode) => ({
  decklink: { slides: { enabled: true, deviceIndex: DEVICE_INDEX, displayMode }, notes: { enabled: false } },
});

async function probeOne(key) {
  const mode = DISPLAY_MODES[key];
  if (!mode) { console.log(`FAIL ${key} (unknown mode)`); return false; }
  const win = fakeWindow(mode.width, mode.height);
  await DecklinkOutputManager.init(() => win, () => null, prefsFor(key));
  const { providerType, slides } = DecklinkOutputManager.getStatus();
  const ok = providerType === 'macadam' && slides.active && !slides.error;
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${key.padEnd(10)} ${ok ? '' : (slides.error || `provider ${providerType}`)}`);
  if (ok) await sleep(HOLD_MS);
  await DecklinkOutputManager.shutdown();
  await sleep(800); // stop() kills the worker after 600ms
  return ok;
}

async function main() {
  if (ONLY) {
    const ok = await probeOne(ONLY);
    process.exit(ok ? 0 : 1);
  }

  const good = [];
  for (const key of Object.keys(DISPLAY_MODES)) {
    const r = spawnSync(process.execPath, [__filename], {
      env: { ...process.env, DECKLINK_ONLY: key },
      stdio: ['ignore', 'pipe', 'pipe'],
      encoding: 'utf8',
    });
    const line = (r.stdout || '').split('\n').find((l) => /^(OK|FAIL)/.test(l)) || `FAIL ${key} (no result)`;
    console.log(line.trimEnd());
    if (r.status === 0) good.push(key);
    await sleep(GAP_MS);
  }
  console.log(`\nSupported on device ${DEVICE_INDEX}: ${good.length ? good.join(', ') : '(none)'}`);
  process.exit(good.length ? 0 : 1);
}

main().catch((e) => { console.error(e); process.exit(2); });
