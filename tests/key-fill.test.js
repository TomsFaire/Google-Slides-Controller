const { test } = require('node:test');
const assert = require('node:assert/strict');

// ── Inlined from main.js (Electron can't be required in test) ─────────────────

function isAllowedKeyFillUrl(urlString) {
  try {
    new URL(String(urlString || '').trim());
    return true;
  } catch (e) {
    return false;
  }
}

function getKeyFillFillWindowOptions(bounds, primaryBounds) {
  const b = bounds && bounds.width ? bounds : primaryBounds;
  return {
    x: b.x,
    y: b.y,
    width: b.width,
    height: b.height,
    frame: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      partition: 'persist:keyfill'
    }
  };
}

function getKeyFillKeyWindowOptions(bounds, primaryBounds) {
  const b = bounds && bounds.width ? bounds : primaryBounds;
  return {
    x: b.x,
    y: b.y,
    width: b.width,
    height: b.height,
    frame: false,
    backgroundColor: '#000000',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      partition: 'persist:keyfill'
    }
  };
}

// Simulates the request-body validation path in POST /api/open-key-fill
function validateKeyFillBody(data) {
  const fillUrl = (data.fillUrl || '').trim();
  const keyUrl = (data.keyUrl || '').trim();
  if (!fillUrl || !keyUrl) return { error: 'fillUrl and keyUrl are required' };
  if (!isAllowedKeyFillUrl(fillUrl)) return { error: 'fillUrl must be a valid URL' };
  if (!isAllowedKeyFillUrl(keyUrl)) return { error: 'keyUrl must be a valid URL' };
  return { fillUrl, keyUrl };
}

// ── URL validation ─────────────────────────────────────────────────────────────

test('isAllowedKeyFillUrl: accepts https URLs', () => {
  assert.equal(isAllowedKeyFillUrl('https://example.com'), true);
  assert.equal(isAllowedKeyFillUrl('https://sub.domain.io/path?q=1'), true);
});

test('isAllowedKeyFillUrl: accepts http URLs', () => {
  assert.equal(isAllowedKeyFillUrl('http://example.com'), true);
  assert.equal(isAllowedKeyFillUrl('http://192.168.1.100:8080/overlay'), true);
});

test('isAllowedKeyFillUrl: rejects non-URL strings', () => {
  assert.equal(isAllowedKeyFillUrl('not-a-url'), false);
  assert.equal(isAllowedKeyFillUrl(''), false);
  assert.equal(isAllowedKeyFillUrl(null), false);
  assert.equal(isAllowedKeyFillUrl(undefined), false);
});

// ── Window options: fill ───────────────────────────────────────────────────────

const PRIMARY = { x: 0, y: 0, width: 1920, height: 1080 };
const SECONDARY = { x: 1920, y: 0, width: 2560, height: 1440 };

test('getKeyFillFillWindowOptions: uses provided bounds', () => {
  const opts = getKeyFillFillWindowOptions(SECONDARY, PRIMARY);
  assert.equal(opts.x, 1920);
  assert.equal(opts.width, 2560);
  assert.equal(opts.height, 1440);
});

test('getKeyFillFillWindowOptions: falls back to primary when bounds missing', () => {
  const opts = getKeyFillFillWindowOptions(null, PRIMARY);
  assert.equal(opts.x, 0);
  assert.equal(opts.width, 1920);
});

test('getKeyFillFillWindowOptions: is frameless with keyfill partition', () => {
  const opts = getKeyFillFillWindowOptions(SECONDARY, PRIMARY);
  assert.equal(opts.frame, false);
  assert.equal(opts.webPreferences.partition, 'persist:keyfill');
  assert.equal(opts.webPreferences.nodeIntegration, false);
  assert.equal(opts.webPreferences.contextIsolation, true);
});

test('getKeyFillFillWindowOptions: does NOT set black backgroundColor (color fill)', () => {
  const opts = getKeyFillFillWindowOptions(SECONDARY, PRIMARY);
  assert.equal(opts.backgroundColor, undefined);
});

// ── Window options: key ────────────────────────────────────────────────────────

test('getKeyFillKeyWindowOptions: uses provided bounds', () => {
  const opts = getKeyFillKeyWindowOptions(SECONDARY, PRIMARY);
  assert.equal(opts.x, 1920);
  assert.equal(opts.width, 2560);
});

test('getKeyFillKeyWindowOptions: has black backgroundColor for luminance key', () => {
  const opts = getKeyFillKeyWindowOptions(SECONDARY, PRIMARY);
  assert.equal(opts.backgroundColor, '#000000');
});

test('getKeyFillKeyWindowOptions: is frameless with keyfill partition', () => {
  const opts = getKeyFillKeyWindowOptions(SECONDARY, PRIMARY);
  assert.equal(opts.frame, false);
  assert.equal(opts.webPreferences.partition, 'persist:keyfill');
});

// ── Key and fill use independent sessions (not Google or Slido partitions) ─────

test('key and fill windows both use persist:keyfill, not persist:google', () => {
  const fill = getKeyFillFillWindowOptions(SECONDARY, PRIMARY);
  const key = getKeyFillKeyWindowOptions(SECONDARY, PRIMARY);
  assert.equal(fill.webPreferences.partition, 'persist:keyfill');
  assert.equal(key.webPreferences.partition, 'persist:keyfill');
  assert.notEqual(fill.webPreferences.partition, 'persist:google');
  assert.notEqual(key.webPreferences.partition, 'persist:slido');
});

// ── API body validation ────────────────────────────────────────────────────────

test('validateKeyFillBody: accepts valid https pair', () => {
  const result = validateKeyFillBody({
    fillUrl: 'https://fill.example.com',
    keyUrl: 'https://key.example.com'
  });
  assert.equal(result.fillUrl, 'https://fill.example.com');
  assert.equal(result.keyUrl, 'https://key.example.com');
  assert.equal(result.error, undefined);
});

test('validateKeyFillBody: rejects missing fillUrl', () => {
  const result = validateKeyFillBody({ keyUrl: 'https://key.example.com' });
  assert.ok(result.error);
  assert.match(result.error, /fillUrl/);
});

test('validateKeyFillBody: rejects missing keyUrl', () => {
  const result = validateKeyFillBody({ fillUrl: 'https://fill.example.com' });
  assert.ok(result.error);
  assert.match(result.error, /keyUrl/);
});

test('validateKeyFillBody: accepts http fillUrl', () => {
  const result = validateKeyFillBody({
    fillUrl: 'http://192.168.1.100:8080/fill',
    keyUrl: 'https://key.example.com'
  });
  assert.equal(result.error, undefined);
});

test('validateKeyFillBody: accepts http keyUrl', () => {
  const result = validateKeyFillBody({
    fillUrl: 'https://fill.example.com',
    keyUrl: 'http://192.168.1.100:8080/key'
  });
  assert.equal(result.error, undefined);
});

test('validateKeyFillBody: rejects non-URL fillUrl', () => {
  const result = validateKeyFillBody({ fillUrl: 'not-a-url', keyUrl: 'https://key.example.com' });
  assert.ok(result.error);
  assert.match(result.error, /fillUrl/);
});

test('validateKeyFillBody: rejects non-URL keyUrl', () => {
  const result = validateKeyFillBody({ fillUrl: 'https://fill.example.com', keyUrl: 'not-a-url' });
  assert.ok(result.error);
  assert.match(result.error, /keyUrl/);
});

test('validateKeyFillBody: trims whitespace from URLs', () => {
  const result = validateKeyFillBody({
    fillUrl: '  https://fill.example.com  ',
    keyUrl: '  https://key.example.com  '
  });
  assert.equal(result.error, undefined);
  assert.equal(result.fillUrl, 'https://fill.example.com');
  assert.equal(result.keyUrl, 'https://key.example.com');
});

test('validateKeyFillBody: fill and key URLs can be different', () => {
  const result = validateKeyFillBody({
    fillUrl: 'https://color-content.example.com/live',
    keyUrl: 'https://matte-source.example.com/key'
  });
  assert.equal(result.error, undefined);
  assert.notEqual(result.fillUrl, result.keyUrl);
});

test('validateKeyFillBody: fill and key URLs can be the same', () => {
  const url = 'https://same.example.com/';
  const result = validateKeyFillBody({ fillUrl: url, keyUrl: url });
  assert.equal(result.error, undefined);
  assert.equal(result.fillUrl, result.keyUrl);
});
