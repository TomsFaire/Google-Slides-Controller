'use strict';

/**
 * Checks what Electron's capturePage() returns for known sRGB colours on each
 * connected display, and prints each display's colour space. Answers: do
 * captured pixels depend on which display the window is on?
 *
 *   yarn decklink:capture-colors
 *
 * If captured values match expected on every display, the DeckLink is fed plain
 * sRGB regardless of the monitor, and a visible difference on one display (e.g.
 * an LG looking pink) comes from that display's profile, not from the capture.
 * If values differ per display, the capture is in the display's space and the
 * app should force sRGB (Chromium --force-color-profile=srgb) or convert.
 *
 * Runs under Electron itself (not ELECTRON_RUN_AS_NODE). Uses a throwaway
 * profile so it never touches your real preferences or Google session.
 */

const os = require('os');
const path = require('path');
const { app, BrowserWindow, screen } = require('electron');

app.setPath('userData', path.join(os.tmpdir(), 'decklink-capture-colors'));

const PATCH_W = 100;
const PATCH_H = 300;
const PATCHES = [
  ['red',        [255, 0, 0]],
  ['green',      [0, 255, 0]],
  ['blue',       [0, 0, 255]],
  ['white',      [255, 255, 255]],
  ['grey 50%',   [128, 128, 128]],
  ['grey 75%',   [191, 191, 191]],
  ['skin',       [200, 150, 120]],
  ['pink-ish',   [255, 150, 200]],
];

const html = `<!doctype html><body style="margin:0;display:flex;background:#000">${PATCHES.map(
  ([, [r, g, b]]) => `<div style="width:${PATCH_W}px;height:${PATCH_H}px;background:rgb(${r},${g},${b})"></div>`
).join('')}</body>`;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function checkDisplay(display) {
  const { x, y, width, height } = display.bounds;
  console.log(`\n== display ${display.id} "${display.label || ''}" ${width}x${height} scale ${display.scaleFactor}`);
  console.log(`   colorSpace: ${JSON.stringify(display.colorSpace)}  depth: ${display.colorDepth}`);

  const win = new BrowserWindow({
    x, y, width: PATCH_W * PATCHES.length, height: PATCH_H,
    frame: false, show: true, useContentSize: true, backgroundColor: '#000',
    webPreferences: { backgroundThrottling: false, sandbox: true },
  });
  await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
  await sleep(1200);

  const image = await win.webContents.capturePage();
  const { width: iw } = image.getSize();
  const scale = iw / (PATCH_W * PATCHES.length);
  const bmp = image.toBitmap(); // BGRA

  let worst = 0;
  for (const [i, [name, [er, eg, eb]]] of PATCHES.entries()) {
    const px = Math.round((i * PATCH_W + PATCH_W / 2) * scale);
    const py = Math.round((PATCH_H / 2) * scale);
    const o = (py * iw + px) * 4;
    const [r, g, b] = [bmp[o + 2], bmp[o + 1], bmp[o]];
    const d = Math.max(Math.abs(r - er), Math.abs(g - eg), Math.abs(b - eb));
    worst = Math.max(worst, d);
    console.log(`   ${name.padEnd(9)} expected ${[er, eg, eb].join(',').padEnd(11)} captured ${[r, g, b].join(',').padEnd(11)} ${d > 2 ? `  <-- off by ${d}` : ''}`);
  }
  console.log(`   => ${worst <= 2 ? 'capture matches sRGB on this display' : `capture DIFFERS from sRGB here (max off by ${worst})`}`);
  win.destroy();
  return worst;
}

app.whenReady().then(async () => {
  const displays = screen.getAllDisplays();
  console.log(`Electron ${process.versions.electron}, ${displays.length} display(s)`);
  const results = [];
  for (const d of displays) results.push([d, await checkDisplay(d)]);

  console.log('\nSummary:');
  for (const [d, worst] of results) {
    console.log(`  display ${d.id} "${d.label || ''}": max deviation ${worst}`);
  }
  app.quit();
}).catch((e) => { console.error(e); app.exit(1); });

app.on('window-all-closed', () => {});
