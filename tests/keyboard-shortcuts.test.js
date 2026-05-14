const { test } = require('node:test');
const assert = require('node:assert/strict');

// ── Inlined from main.js (browser globals shimmed for test) ──────────────────

function getShortcutModifier(platformString) {
  if (platformString && platformString.__isUserAgentData) {
    return platformString.platform === 'macOS' ? '⌘' : 'Ctrl';
  }
  return /Mac|iPhone|iPad|iPod/.test(platformString) ? '⌘' : 'Ctrl';
}

// ── Tests ────────────────────────────────────────────────────────────────────

test('getShortcutModifier returns ⌘ for MacIntel platform string', () => {
  assert.equal(getShortcutModifier('MacIntel'), '⌘');
});

test('getShortcutModifier returns ⌘ for iPhone', () => {
  assert.equal(getShortcutModifier('iPhone'), '⌘');
});

test('getShortcutModifier returns ⌘ for iPad', () => {
  assert.equal(getShortcutModifier('iPad'), '⌘');
});

test('getShortcutModifier returns Ctrl for Win32', () => {
  assert.equal(getShortcutModifier('Win32'), 'Ctrl');
});

test('getShortcutModifier returns Ctrl for Linux x86_64', () => {
  assert.equal(getShortcutModifier('Linux x86_64'), 'Ctrl');
});

test('getShortcutModifier returns Ctrl for empty string', () => {
  assert.equal(getShortcutModifier(''), 'Ctrl');
});
