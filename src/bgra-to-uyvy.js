'use strict';

/**
 * Converts a BGRA frame to 8-bit 4:2:2 UYVY (DeckLink bmdFormat8BitYUV),
 * BT.709 limited range. Chroma is averaged over each horizontal pixel pair.
 * Some DeckLink modes (e.g. 1080p50/59.94/60 on the UltraStudio Monitor 3G)
 * accept only YUV, not BGRA.
 */
function bgraToUyvy(bgra, width, height, out) {
  const pairs = (width * height) >> 1;
  const dst = out || Buffer.allocUnsafe(width * height * 2);

  for (let i = 0, s = 0, d = 0; i < pairs; i++, s += 8, d += 4) {
    const b0 = bgra[s], g0 = bgra[s + 1], r0 = bgra[s + 2];
    const b1 = bgra[s + 4], g1 = bgra[s + 5], r1 = bgra[s + 6];
    const r = (r0 + r1) >> 1, g = (g0 + g1) >> 1, b = (b0 + b1) >> 1;

    dst[d]     = ((-26 * r - 87 * g + 112 * b + 128) >> 8) + 128;           // U
    dst[d + 1] = ((47 * r0 + 157 * g0 + 16 * b0 + 128) >> 8) + 16;          // Y0
    dst[d + 2] = ((112 * r - 102 * g - 10 * b + 128) >> 8) + 128;           // V
    dst[d + 3] = ((47 * r1 + 157 * g1 + 16 * b1 + 128) >> 8) + 16;          // Y1
  }
  return dst;
}

module.exports = { bgraToUyvy };
