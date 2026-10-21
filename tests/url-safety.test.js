const { test } = require('node:test');
const assert = require('node:assert/strict');

const { isSafeWindowOpenUrl, isExternalHttpUrl } = require('../src/url-safety');

test('window.open allows https, http and about:blank', () => {
  assert.equal(isSafeWindowOpenUrl('https://docs.google.com/presentation/d/x/present'), true);
  assert.equal(isSafeWindowOpenUrl('http://localhost:9595/'), true);
  assert.equal(isSafeWindowOpenUrl('about:blank'), true);
});

test('window.open refuses file, javascript, data, custom schemes and junk', () => {
  for (const u of ['file:///etc/passwd', 'javascript:alert(1)', 'data:text/html,hi', 'foo://bar', 'not a url', '']) {
    assert.equal(isSafeWindowOpenUrl(u), false, u);
  }
});

test('isExternalHttpUrl only accepts http(s)', () => {
  assert.equal(isExternalHttpUrl('https://example.com'), true);
  assert.equal(isExternalHttpUrl('about:blank'), false);
  assert.equal(isExternalHttpUrl('file:///x'), false);
});
