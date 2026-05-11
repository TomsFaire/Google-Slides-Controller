const { test } = require('node:test');
const assert = require('node:assert/strict');
const { normalizePerfectCuePorts } = require('../src/perfectcue-port-config');

test('plain number array migrates to PortConfig objects with dsan adapter', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [8899, 18899] });
  assert.deepEqual(result, [
    { port: 8899, name: '', enabled: true, adapter: 'dsan' },
    { port: 18899, name: '', enabled: true, adapter: 'dsan' }
  ]);
});

test('existing PortConfig objects pass through with adapter', () => {
  const result = normalizePerfectCuePorts({
    perfectCuePorts: [{ port: 8899, name: 'Extender 1', enabled: false, adapter: 'waveshare' }]
  });
  assert.deepEqual(result, [{ port: 8899, name: 'Extender 1', enabled: false, adapter: 'waveshare' }]);
});

test('adapter defaults to dsan when missing', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [{ port: 8899, name: 'Test' }] });
  assert.equal(result[0].adapter, 'dsan');
});

test('invalid adapter value normalizes to dsan', () => {
  const result = normalizePerfectCuePorts({
    perfectCuePorts: [{ port: 8899, name: '', adapter: 'unknown' }]
  });
  assert.equal(result[0].adapter, 'dsan');
});

test('enabled field defaults to true when missing', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [{ port: 8899, name: 'Test' }] });
  assert.equal(result[0].enabled, true);
});

test('empty array falls back to default port 8899 with dsan', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [] });
  assert.deepEqual(result, [{ port: 8899, name: '', enabled: true, adapter: 'dsan' }]);
});

test('legacy perfectCuePort (single number) is used when array is empty', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [], perfectCuePort: 9999 });
  assert.deepEqual(result, [{ port: 9999, name: '', enabled: true, adapter: 'dsan' }]);
});

test('invalid port numbers (<=0) are filtered out', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [0, 8899, -1] });
  assert.deepEqual(result, [{ port: 8899, name: '', enabled: true, adapter: 'dsan' }]);
});

test('mixed number and object array migrates correctly', () => {
  const result = normalizePerfectCuePorts({
    perfectCuePorts: [8899, { port: 18899, name: 'Kit 2', enabled: false, adapter: 'waveshare' }]
  });
  assert.deepEqual(result, [
    { port: 8899, name: '', enabled: true, adapter: 'dsan' },
    { port: 18899, name: 'Kit 2', enabled: false, adapter: 'waveshare' }
  ]);
});

test('no perfectCuePorts key falls back to default port 8899', () => {
  const result = normalizePerfectCuePorts({});
  assert.deepEqual(result, [{ port: 8899, name: '', enabled: true, adapter: 'dsan' }]);
});
