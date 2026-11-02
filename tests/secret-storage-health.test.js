const { test } = require('node:test');
const assert = require('node:assert/strict');
const {
  LEVEL_OS,
  LEVEL_WEAK,
  LEVEL_NONE,
  assessSecretStorage
} = require('../src/secret-storage-health');

test('an OS keystore counts as protected', () => {
  const r = assessSecretStorage({ available: true, backend: 'gnome-libsecret', platform: 'linux' });
  assert.equal(r.level, LEVEL_OS);
  assert.equal(r.protected, true);
  assert.match(r.detail, /gnome-libsecret/);
});

test('macOS reports protected with no backend name', () => {
  // getSelectedStorageBackend() is Linux-only, so backend is undefined here.
  const r = assessSecretStorage({ available: true, backend: undefined, platform: 'darwin' });
  assert.equal(r.level, LEVEL_OS);
  assert.equal(r.protected, true);
  assert.equal(r.backend, null);
});

test('basic_text is reported as not protected even though encryption "works"', () => {
  // The case this whole module exists for: Electron returns a base64 blob that
  // looks encrypted, but the key is a hardcoded password.
  const r = assessSecretStorage({ available: true, backend: 'basic_text', platform: 'linux' });
  assert.equal(r.level, LEVEL_WEAK);
  assert.equal(r.protected, false);
  assert.match(r.detail, /hardcoded password/);
});

test('unavailable encryption is reported as plaintext', () => {
  const r = assessSecretStorage({ available: false, platform: 'linux' });
  assert.equal(r.level, LEVEL_NONE);
  assert.equal(r.protected, false);
  assert.match(r.detail, /plain text/);
});

test('the unavailable hint is tailored per platform', () => {
  const linux = assessSecretStorage({ available: false, platform: 'linux' });
  const mac = assessSecretStorage({ available: false, platform: 'darwin' });
  assert.match(linux.detail, /gnome-keyring|kwallet/);
  assert.doesNotMatch(mac.detail, /gnome-keyring/);
});

test('unavailable wins over a reported backend', () => {
  const r = assessSecretStorage({ available: false, backend: 'gnome-libsecret', platform: 'linux' });
  assert.equal(r.level, LEVEL_NONE);
  assert.equal(r.protected, false);
});

test('an empty backend string is normalised to null', () => {
  assert.equal(assessSecretStorage({ available: true, backend: '' }).backend, null);
});

test('missing input does not throw and is treated as unprotected', () => {
  const r = assessSecretStorage();
  assert.equal(r.protected, false);
  assert.equal(r.level, LEVEL_NONE);
});

test('every result carries a short summary for the UI', () => {
  for (const input of [
    { available: true, backend: 'kwallet6' },
    { available: true, backend: 'basic_text' },
    { available: false }
  ]) {
    const r = assessSecretStorage(input);
    assert.equal(typeof r.summary, 'string');
    assert.ok(r.summary.length > 0 && r.summary.length < 40, `bad summary: ${r.summary}`);
  }
});
