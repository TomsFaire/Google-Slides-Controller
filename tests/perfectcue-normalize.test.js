const { test } = require('node:test');
const assert = require('node:assert/strict');

// Inline copy of normalizePerfectCuePorts for unit testing (main.js requires Electron)
function normalizePerfectCuePorts(prefs) {
  const raw = Array.isArray(prefs.perfectCuePorts) ? prefs.perfectCuePorts : [];
  const configs = raw.map(entry => {
    if (typeof entry === 'number') {
      return { port: entry, name: '', enabled: true };
    }
    // Already an object — ensure all three fields are present
    return {
      port: Number(entry.port),
      name: typeof entry.name === 'string' ? entry.name : '',
      enabled: entry.enabled !== false
    };
  }).filter(c => c.port > 0);

  if (configs.length === 0) {
    // Fall back to legacy single-port pref, then hard default
    const legacyPort = prefs.perfectCuePort ? Number(prefs.perfectCuePort) : 0;
    return [{ port: legacyPort > 0 ? legacyPort : 8899, name: '', enabled: true }];
  }
  return configs;
}

test('plain number array migrates to PortConfig objects', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [8899, 18899] });
  assert.deepEqual(result, [
    { port: 8899, name: '', enabled: true },
    { port: 18899, name: '', enabled: true }
  ]);
});

test('existing PortConfig objects pass through unchanged', () => {
  const result = normalizePerfectCuePorts({
    perfectCuePorts: [{ port: 8899, name: 'Extender 1', enabled: false }]
  });
  assert.deepEqual(result, [{ port: 8899, name: 'Extender 1', enabled: false }]);
});

test('enabled field defaults to true when missing', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [{ port: 8899, name: 'Test' }] });
  assert.equal(result[0].enabled, true);
});

test('empty array falls back to default port 8899', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [] });
  assert.deepEqual(result, [{ port: 8899, name: '', enabled: true }]);
});

test('legacy perfectCuePort (single number) is used when array is empty', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [], perfectCuePort: 9999 });
  assert.deepEqual(result, [{ port: 9999, name: '', enabled: true }]);
});

test('invalid port numbers (<=0) are filtered out', () => {
  const result = normalizePerfectCuePorts({ perfectCuePorts: [0, 8899, -1] });
  assert.deepEqual(result, [{ port: 8899, name: '', enabled: true }]);
});

test('mixed number and object array migrates correctly', () => {
  const result = normalizePerfectCuePorts({
    perfectCuePorts: [8899, { port: 18899, name: 'Kit 2', enabled: false }]
  });
  assert.deepEqual(result, [
    { port: 8899, name: '', enabled: true },
    { port: 18899, name: 'Kit 2', enabled: false }
  ]);
});

test('no perfectCuePorts key falls back to default port 8899', () => {
  const result = normalizePerfectCuePorts({});
  assert.deepEqual(result, [{ port: 8899, name: '', enabled: true }]);
});
