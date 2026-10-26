const { test } = require('node:test');
const assert = require('node:assert/strict');
const { bgraToUyvy } = require('../src/bgra-to-uyvy');

// Two-pixel BGRA frame -> one UYVY macropixel [U, Y0, V, Y1].
const frame = (r, g, b) => Buffer.from([b, g, r, 255, b, g, r, 255]);
const near = (actual, expected, tol = 2) =>
  assert.ok(Math.abs(actual - expected) <= tol, `${actual} not within ${tol} of ${expected}`);

test('black maps to Y=16, U=V=128', () => {
  const [u, y0, v, y1] = bgraToUyvy(frame(0, 0, 0), 2, 1);
  assert.deepEqual([u, y0, v, y1], [128, 16, 128, 16]);
});

test('white maps to Y=235, U=V=128', () => {
  const [u, y0, v, y1] = bgraToUyvy(frame(255, 255, 255), 2, 1);
  near(y0, 235); near(y1, 235); near(u, 128); near(v, 128);
});

test('pure red has high V and low U', () => {
  const [u, y0, v] = bgraToUyvy(frame(255, 0, 0), 2, 1);
  near(y0, 63); near(u, 102); near(v, 240);
});

test('pure blue has high U and low V', () => {
  const [u, y0, v] = bgraToUyvy(frame(0, 0, 255), 2, 1);
  near(y0, 32); near(u, 240); near(v, 118);
});

test('output is width*height*2 bytes and reuses a supplied buffer', () => {
  const w = 8, h = 4;
  const src = Buffer.alloc(w * h * 4);
  assert.equal(bgraToUyvy(src, w, h).length, w * h * 2);
  const out = Buffer.alloc(w * h * 2);
  assert.equal(bgraToUyvy(src, w, h, out), out);
});

test('1080p frame converts fast enough for 60fps budget', () => {
  const w = 1920, h = 1080;
  const src = Buffer.alloc(w * h * 4, 0x80);
  const out = Buffer.alloc(w * h * 2);
  bgraToUyvy(src, w, h, out); // warm up
  const t0 = process.hrtime.bigint();
  for (let i = 0; i < 10; i++) bgraToUyvy(src, w, h, out);
  const msPerFrame = Number(process.hrtime.bigint() - t0) / 1e6 / 10;
  console.log(`  bgraToUyvy 1080p: ${msPerFrame.toFixed(1)} ms/frame`);
  assert.ok(msPerFrame < 16, `${msPerFrame.toFixed(1)}ms exceeds a 60fps frame budget`);
});

test('every matrix/range keeps full-scale inputs inside 0-255 without wrapping', () => {
  const { COEFFS } = require('../src/bgra-to-uyvy');
  const inputs = [[0,0,0],[255,255,255],[255,0,0],[0,255,0],[0,0,255],[255,255,0],[0,255,255],[255,0,255]];
  for (const key of Object.keys(COEFFS)) {
    const [matrix, range] = key.split('-');
    for (const [r, g, b] of inputs) {
      const out = bgraToUyvy(frame(r, g, b), 2, 1, undefined, { matrix, range });
      const lo = range === 'limited' ? 16 : 0;
      const hi = range === 'limited' ? 240 : 255;
      for (const [i, v] of [...out].entries()) {
        assert.ok(v >= lo - 1 && v <= hi + 1, `${key} rgb(${r},${g},${b}) byte ${i} = ${v}`);
      }
    }
  }
});

test('full range keeps black at 0 and white at ~255; 601 differs from 709 on saturated colours', () => {
  const black = bgraToUyvy(frame(0, 0, 0), 2, 1, undefined, { range: 'full' });
  const white = bgraToUyvy(frame(255, 255, 255), 2, 1, undefined, { range: 'full' });
  assert.equal(black[1], 0);
  near(white[1], 255);
  const g709 = bgraToUyvy(frame(0, 255, 0), 2, 1, undefined, { matrix: '709' });
  const g601 = bgraToUyvy(frame(0, 255, 0), 2, 1, undefined, { matrix: '601' });
  assert.notEqual(g709[1], g601[1]);
});

test('unknown matrix/range is rejected', () => {
  assert.throws(() => bgraToUyvy(frame(0, 0, 0), 2, 1, undefined, { matrix: '2020' }), /unsupported/);
});
