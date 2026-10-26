'use strict';

/**
 * Converts a BGRA frame to 8-bit 4:2:2 UYVY (DeckLink bmdFormat8BitYUV).
 * Chroma is averaged over each horizontal pixel pair.
 * Some DeckLink modes (e.g. 1080p50/59.94/60 on the UltraStudio Monitor 3G)
 * accept only YUV, not BGRA.
 *
 * Defaults are BT.709 limited range, which is what HD SDI/HDMI monitors
 * expect. matrix/range are selectable only so colour problems can be A/B
 * tested (see scripts/decklink-color-chart.js).
 *
 * Coefficients are integer approximations scaled by 256.
 */
const COEFFS = {
  '709-limited': { y: [47, 157, 16],  yOff: 16, cb: [-26, -87, 112],  cr: [112, -102, -10] },
  '601-limited': { y: [66, 129, 25],  yOff: 16, cb: [-38, -74, 112],  cr: [112, -94, -18] },
  // Full-range sums are kept <= 255 so 255-valued inputs cannot wrap in a Buffer.
  '709-full':    { y: [54, 183, 18],  yOff: 0,  cb: [-29, -98, 127],  cr: [127, -115, -12] },
  '601-full':    { y: [77, 150, 29],  yOff: 0,  cb: [-43, -84, 127],  cr: [127, -106, -21] },
};

function coeffsFor(matrix = '709', range = 'limited') {
  const c = COEFFS[`${matrix}-${range}`];
  if (!c) throw new Error(`unsupported YUV conversion: matrix=${matrix} range=${range}`);
  return c;
}

function bgraToUyvy(bgra, width, height, out, { matrix = '709', range = 'limited' } = {}) {
  const { y: ky, yOff, cb: kb, cr: kr } = coeffsFor(matrix, range);
  const pairs = (width * height) >> 1;
  const dst = out || Buffer.allocUnsafe(width * height * 2);

  for (let i = 0, s = 0, d = 0; i < pairs; i++, s += 8, d += 4) {
    const b0 = bgra[s], g0 = bgra[s + 1], r0 = bgra[s + 2];
    const b1 = bgra[s + 4], g1 = bgra[s + 5], r1 = bgra[s + 6];
    const r = (r0 + r1) >> 1, g = (g0 + g1) >> 1, b = (b0 + b1) >> 1;

    dst[d]     = ((kb[0] * r + kb[1] * g + kb[2] * b + 128) >> 8) + 128;                 // U
    dst[d + 1] = ((ky[0] * r0 + ky[1] * g0 + ky[2] * b0 + 128) >> 8) + yOff;             // Y0
    dst[d + 2] = ((kr[0] * r + kr[1] * g + kr[2] * b + 128) >> 8) + 128;                 // V
    dst[d + 3] = ((ky[0] * r1 + ky[1] * g1 + ky[2] * b1 + 128) >> 8) + yOff;             // Y1
  }
  return dst;
}

module.exports = { bgraToUyvy, COEFFS };
