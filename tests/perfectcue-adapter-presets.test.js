const { test } = require('node:test');
const assert = require('node:assert/strict');
const { normalizeAdapterId, getPerfectCueAdapterPreset } = require('../src/perfectcue-adapter-presets');

test('normalizeAdapterId waveshare', () => {
  assert.equal(normalizeAdapterId('waveshare'), 'waveshare');
});

test('normalizeAdapterId defaults unknown to dsan', () => {
  assert.equal(normalizeAdapterId(undefined), 'dsan');
  assert.equal(normalizeAdapterId(''), 'dsan');
  assert.equal(normalizeAdapterId('other'), 'dsan');
});

test('getPerfectCueAdapterPreset dsan shorter ping than waveshare', () => {
  const d = getPerfectCueAdapterPreset('dsan');
  const w = getPerfectCueAdapterPreset('waveshare');
  assert.ok(d.pingIntervalMs < w.pingIntervalMs);
  assert.ok(d.idleTimeoutMs < w.idleTimeoutMs);
});
