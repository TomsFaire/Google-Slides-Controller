const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {
  SECRET_FILE_MODE,
  isModeTooPermissive,
  writeSecretFileSync,
  ensureSecretFileMode
} = require('../src/secret-files');

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'gso-secret-'));
const modeOf = (p) => fs.statSync(p).mode & 0o777;

test('isModeTooPermissive flags any group or other bit', () => {
  assert.equal(isModeTooPermissive(0o600), false);
  assert.equal(isModeTooPermissive(0o400), false);
  assert.equal(isModeTooPermissive(0o640), true);
  assert.equal(isModeTooPermissive(0o604), true);
  assert.equal(isModeTooPermissive(0o644), true);
  assert.equal(isModeTooPermissive(0o666), true);
});

test('isModeTooPermissive ignores the file-type bits in st_mode', () => {
  // st_mode from statSync carries S_IFREG (0o100000) above the permission bits.
  assert.equal(isModeTooPermissive(0o100600), false);
  assert.equal(isModeTooPermissive(0o100644), true);
});

test('writeSecretFileSync creates a new file owner-only', () => {
  const dir = tmp();
  const p = path.join(dir, 'prefs.json');
  writeSecretFileSync(p, '{"a":1}', 'utf8');
  assert.equal(modeOf(p), SECRET_FILE_MODE);
  assert.equal(fs.readFileSync(p, 'utf8'), '{"a":1}');
});

test('writeSecretFileSync tightens a file that already exists world-readable', () => {
  // The regression that matters: writeFileSync's mode option is ignored when
  // the file already exists, so without the chmod an upgraded install would
  // keep its old 0644 preferences file forever.
  const dir = tmp();
  const p = path.join(dir, 'prefs.json');
  fs.writeFileSync(p, 'old', { mode: 0o644 });
  fs.chmodSync(p, 0o644);
  assert.equal(modeOf(p), 0o644);

  writeSecretFileSync(p, 'new', 'utf8');
  assert.equal(modeOf(p), SECRET_FILE_MODE);
  assert.equal(fs.readFileSync(p, 'utf8'), 'new');
});

test('writeSecretFileSync handles Buffer payloads with no encoding', () => {
  const dir = tmp();
  const p = path.join(dir, 'secret.bin');
  const buf = Buffer.from([1, 2, 3, 4]);
  writeSecretFileSync(p, buf);
  assert.equal(modeOf(p), SECRET_FILE_MODE);
  assert.deepEqual(fs.readFileSync(p), buf);
});

test('ensureSecretFileMode tightens an existing permissive file', () => {
  const dir = tmp();
  const p = path.join(dir, 'cf-tunnel-abc.json');
  fs.writeFileSync(p, '{}');
  fs.chmodSync(p, 0o644);
  assert.equal(ensureSecretFileMode(p), 'tightened');
  assert.equal(modeOf(p), SECRET_FILE_MODE);
});

test('ensureSecretFileMode leaves an already-tight file alone', () => {
  const dir = tmp();
  const p = path.join(dir, 'prefs.json');
  fs.writeFileSync(p, '{}');
  fs.chmodSync(p, 0o600);
  assert.equal(ensureSecretFileMode(p), 'already-owner-only');
  assert.equal(modeOf(p), SECRET_FILE_MODE);
});

test('ensureSecretFileMode reports a missing file rather than throwing', () => {
  const dir = tmp();
  assert.equal(ensureSecretFileMode(path.join(dir, 'nope.json')), 'missing');
});

test('ensureSecretFileMode does not alter file contents', () => {
  const dir = tmp();
  const p = path.join(dir, 'prefs.json');
  fs.writeFileSync(p, '{"token":"abc"}');
  fs.chmodSync(p, 0o666);
  ensureSecretFileMode(p);
  assert.equal(fs.readFileSync(p, 'utf8'), '{"token":"abc"}');
});
