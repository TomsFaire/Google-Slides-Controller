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

test('0x06 mis-frame variant returns next', () => {
  assert.equal(parsePerfectCueByte(0x06), 'next');
});

test('0x86 mis-frame+bit-7 variant returns next', () => {
  assert.equal(parsePerfectCueByte(0x86), 'next');
});

test('0xc6 mis-frame+bits-7+6 variant returns next', () => {
  assert.equal(parsePerfectCueByte(0xc6), 'next');
});

test('0x8c bit-7 noise variant returns next', () => {
  assert.equal(parsePerfectCueByte(0x8c), 'next');
});

test('0xcc bits-7+6 noise variant returns next', () => {
  assert.equal(parsePerfectCueByte(0xcc), 'next');
});

test('0x88 bit-7 noise variant returns previous', () => {
  assert.equal(parsePerfectCueByte(0x88), 'previous');
});

test('0xc8 bits-7+6 noise variant returns previous', () => {
  assert.equal(parsePerfectCueByte(0xc8), 'previous');
});

test('0x84 bit-7 noise variant returns blackout', () => {
  assert.equal(parsePerfectCueByte(0x84), 'blackout');
});

test('0xc4 bits-7+6 noise variant returns blackout', () => {
  assert.equal(parsePerfectCueByte(0xc4), 'blackout');
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
