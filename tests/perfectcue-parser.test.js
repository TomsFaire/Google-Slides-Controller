const { test } = require('node:test');
const assert = require('node:assert/strict');
const { parsePerfectCueByte } = require('../src/perfectcue-parser');

test('0x0F returns next', () => {
  assert.equal(parsePerfectCueByte(0x0F), 'next');
});

test('0x1F returns previous', () => {
  assert.equal(parsePerfectCueByte(0x1F), 'previous');
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
