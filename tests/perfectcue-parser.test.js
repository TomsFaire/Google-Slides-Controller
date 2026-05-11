const { test } = require('node:test');
const assert = require('node:assert/strict');
const { parsePerfectCueByte } = require('../src/perfectcue-parser');

test('0x0c returns next', () => {
  assert.equal(parsePerfectCueByte(0x0c), 'next');
});

test('0x08 returns previous', () => {
  assert.equal(parsePerfectCueByte(0x08), 'previous');
});

test('0x04 returns blackout', () => {
  assert.equal(parsePerfectCueByte(0x04), 'blackout');
});

test('0x8c high-bit noise variant returns next', () => {
  assert.equal(parsePerfectCueByte(0x8c), 'next');
});

test('0x88 high-bit noise variant returns previous', () => {
  assert.equal(parsePerfectCueByte(0x88), 'previous');
});

test('0x84 high-bit noise variant returns blackout', () => {
  assert.equal(parsePerfectCueByte(0x84), 'blackout');
});

test('0xFF keepalive returns null', () => {
  assert.equal(parsePerfectCueByte(0xFF), null);
});

test('0x00 unknown byte returns null', () => {
  assert.equal(parsePerfectCueByte(0x00), null);
});

test('arbitrary unknown byte returns null', () => {
  assert.equal(parsePerfectCueByte(0xAB), null);
});
