'use strict';

/**
 * Outputs a reference colour chart with exact known RGB values so you can
 * check the signal on a monitor's waveform / vectorscope (e.g. Blackmagic
 * Video Assist 12G scopes).
 *
 *   yarn decklink:color-chart                          # 1080p60 (YUV path)
 *   DECKLINK_MODE=1080p2997 yarn decklink:color-chart  # BGRA path (driver converts)
 *   DECKLINK_MODE=1080p2997 DECKLINK_PIXEL_FORMAT=yuv yarn decklink:color-chart
 *   DECKLINK_YUV_MATRIX=601 yarn decklink:color-chart  # A/B the matrix
 *   DECKLINK_YUV_RANGE=full yarn decklink:color-chart  # A/B the range
 *
 * DECKLINK_PIXEL_FORMAT=yuv|bgra forces the path for modes that support both,
 * so the app's conversion can be compared against the driver's. The matrix and
 * range overrides only affect the YUV path. Ctrl+C to stop.
 *
 * Chart layout (top to bottom):
 *   - 75% colour bars   (RGB 191): white yellow cyan green magenta red blue
 *   - 100% colour bars  (RGB 255)
 *   - 11-step greyscale, 0% to 100%
 *
 * Runs under ELECTRON_RUN_AS_NODE (see package.json).
 */

const { DecklinkOutputManager, DISPLAY_MODES } = require('../src/decklink-output');

const MODE = process.env.DECKLINK_MODE || '1080p60';
const DEVICE_INDEX = Number(process.env.DECKLINK_DEVICE || 0);
const FORCED_FORMAT = process.env.DECKLINK_PIXEL_FORMAT;

const BARS = [
  [1, 1, 1], [1, 1, 0], [0, 1, 1], [0, 1, 0], [1, 0, 1], [1, 0, 0], [0, 0, 1],
];

function chart(width, height) {
  const buf = Buffer.alloc(width * height * 4);
  const rows = [
    { to: Math.round(height * 0.55), color: (x) => BARS[Math.floor((x / width) * BARS.length)].map((v) => v * 191) },
    { to: Math.round(height * 0.70), color: (x) => BARS[Math.floor((x / width) * BARS.length)].map((v) => v * 255) },
    { to: height, color: (x) => { const g = Math.round(Math.floor((x / width) * 11) * 25.5); return [g, g, g]; } },
  ];
  let y = 0;
  for (const row of rows) {
    for (; y < row.to; y++) {
      for (let x = 0; x < width; x++) {
        const [r, g, b] = row.color(x);
        const o = (y * width + x) * 4;
        buf[o] = b; buf[o + 1] = g; buf[o + 2] = r; buf[o + 3] = 255;
      }
    }
  }
  return buf;
}

async function main() {
  const mode = DISPLAY_MODES[MODE];
  if (!mode) throw new Error(`unknown DECKLINK_MODE ${MODE}`);
  if (FORCED_FORMAT) mode.pixelFormat = FORCED_FORMAT; // shared object read by the provider
  const path = (mode.pixelFormat || 'bgra') === 'yuv' ? 'YUV (app converts)' : 'BGRA (driver converts)';

  const bitmap = chart(mode.width, mode.height);
  const image = { getSize: () => ({ width: mode.width, height: mode.height }), resize: () => image, toBitmap: () => bitmap };
  const win = { isDestroyed: () => false, webContents: { capturePage: async () => image } };

  await DecklinkOutputManager.init(() => win, () => null, {
    decklink: { slides: { enabled: true, deviceIndex: DEVICE_INDEX, displayMode: MODE }, notes: { enabled: false } },
  });
  const { providerType, slides } = DecklinkOutputManager.getStatus();
  if (providerType !== 'macadam' || slides.error) {
    console.error(`could not start output: provider=${providerType} error=${slides.error}`);
    process.exit(1);
  }

  console.log(`Chart on device ${DEVICE_INDEX}: ${MODE}, ${path}`);
  if ((mode.pixelFormat || 'bgra') === 'yuv') {
    console.log(`YUV matrix=${process.env.DECKLINK_YUV_MATRIX || '709'} range=${process.env.DECKLINK_YUV_RANGE || 'limited'}`);
  }
  console.log(`
Check on the monitor's scopes (waveform + vectorscope):
  - 75% bars: white sits at ~75%; vectorscope dots land in the 75% target boxes
  - 100% bars: white at ~100%
  - greyscale: evenly spaced steps, black at 0%, white at 100%
  Everything squashed (black ~6%, white ~92%)   -> range mismatch (limited vs full)
  Dots near but outside the boxes, hue rotated  -> matrix mismatch (709 vs 601)
  Looks right here but wrong in the app         -> the problem is in what capturePage returns
Ctrl+C to stop.`);

  const stop = async () => { await DecklinkOutputManager.shutdown(); setTimeout(() => process.exit(0), 800); };
  process.on('SIGINT', stop);
  process.on('SIGTERM', stop);
  setInterval(() => {}, 1 << 30);
}

main().catch((e) => { console.error(e); process.exit(2); });
