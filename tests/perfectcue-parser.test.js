const { test } = require('node:test');
const assert = require('node:assert/strict');
const { parsePerfectCueCommand, splitOnCR } = require('../src/perfectcue-parser');

test('parsePerfectCueCommand: >FORWARD returns next', () => {
  assert.equal(parsePerfectCueCommand('>FORWARD 35'), 'next');
});

test('parsePerfectCueCommand: >REVERSE returns previous', () => {
  assert.equal(parsePerfectCueCommand('>REVERSE 3C'), 'previous');
});

test('parsePerfectCueCommand: bare >FORWARD returns next', () => {
  assert.equal(parsePerfectCueCommand('>FORWARD'), 'next');
});

test('parsePerfectCueCommand: unknown line returns null', () => {
  assert.equal(parsePerfectCueCommand('>STATUS OK'), null);
});

test('parsePerfectCueCommand: empty string returns null', () => {
  assert.equal(parsePerfectCueCommand(''), null);
});

test('splitOnCR: splits on carriage return, returns remainder', () => {
  const result = splitOnCR('>FORWARD 35\r>REVERSE 3C\r');
  assert.deepEqual(result.lines, ['>FORWARD 35', '>REVERSE 3C']);
  assert.equal(result.remainder, '');
});

test('splitOnCR: keeps incomplete last fragment as remainder', () => {
  const result = splitOnCR('>FORWARD 35\r>REVER');
  assert.deepEqual(result.lines, ['>FORWARD 35']);
  assert.equal(result.remainder, '>REVER');
});

test('splitOnCR: no CR returns empty lines and full string as remainder', () => {
  const result = splitOnCR('>FORWARD 35');
  assert.deepEqual(result.lines, []);
  assert.equal(result.remainder, '>FORWARD 35');
});

test('splitOnCR: filters out empty lines between consecutive CRs', () => {
  const result = splitOnCR('>FORWARD 35\r\r>REVERSE 3C\r');
  assert.deepEqual(result.lines, ['>FORWARD 35', '>REVERSE 3C']);
});
