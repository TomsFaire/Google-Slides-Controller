const { test } = require('node:test');
const assert = require('node:assert/strict');

// ── Inlined from main.js (browser globals shimmed for test) ──────────────────

function getShortcutModifier(platformString) {
  return /Mac|iPhone|iPad|iPod/.test(platformString) ? '⌘' : 'Ctrl';
}

const KEYBOARD_PRESETS = {
  'cmd+arrow':       e => (e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey,
  'alt+arrow':       e => e.altKey && !e.metaKey && !e.ctrlKey && !e.shiftKey,
  'cmd+shift+arrow': e => (e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey,
};

function getPresetHintText(preset, platformString) {
  const mod = getShortcutModifier(platformString);
  const isMac = mod === '⌘';
  if (preset === 'alt+arrow') {
    const alt = isMac ? '⌥' : 'Alt+';
    return alt + '← Prev slide · ' + alt + '→ Next slide · ' + alt + '↑ Notes up · ' + alt + '↓ Notes down';
  }
  if (preset === 'cmd+shift+arrow') {
    const combo = isMac ? '⌘⇧' : 'Ctrl+Shift+';
    return combo + '← Prev slide · ' + combo + '→ Next slide · ' + combo + '↑ Notes up · ' + combo + '↓ Notes down';
  }
  return mod + '← Prev slide · ' + mod + '→ Next slide · ' + mod + '↑ Notes up · ' + mod + '↓ Notes down';
}

// ── getShortcutModifier ───────────────────────────────────────────────────────

test('getShortcutModifier returns ⌘ for Mac', () => {
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

// ── KEYBOARD_PRESETS ──────────────────────────────────────────────────────────

test('cmd+arrow: matches metaKey alone', () => {
  assert.ok(KEYBOARD_PRESETS['cmd+arrow']({ metaKey: true, ctrlKey: false, shiftKey: false, altKey: false }));
});

test('cmd+arrow: matches ctrlKey alone', () => {
  assert.ok(KEYBOARD_PRESETS['cmd+arrow']({ metaKey: false, ctrlKey: true, shiftKey: false, altKey: false }));
});

test('cmd+arrow: rejects altKey', () => {
  assert.ok(!KEYBOARD_PRESETS['cmd+arrow']({ metaKey: true, ctrlKey: false, shiftKey: false, altKey: true }));
});

test('cmd+arrow: rejects shiftKey', () => {
  assert.ok(!KEYBOARD_PRESETS['cmd+arrow']({ metaKey: true, ctrlKey: false, shiftKey: true, altKey: false }));
});

test('cmd+arrow: rejects no modifier', () => {
  assert.ok(!KEYBOARD_PRESETS['cmd+arrow']({ metaKey: false, ctrlKey: false, shiftKey: false, altKey: false }));
});

test('alt+arrow: matches altKey alone', () => {
  assert.ok(KEYBOARD_PRESETS['alt+arrow']({ altKey: true, metaKey: false, ctrlKey: false, shiftKey: false }));
});

test('alt+arrow: rejects metaKey', () => {
  assert.ok(!KEYBOARD_PRESETS['alt+arrow']({ altKey: true, metaKey: true, ctrlKey: false, shiftKey: false }));
});

test('alt+arrow: rejects ctrlKey', () => {
  assert.ok(!KEYBOARD_PRESETS['alt+arrow']({ altKey: true, metaKey: false, ctrlKey: true, shiftKey: false }));
});

test('alt+arrow: rejects no modifier', () => {
  assert.ok(!KEYBOARD_PRESETS['alt+arrow']({ altKey: false, metaKey: false, ctrlKey: false, shiftKey: false }));
});

test('cmd+shift+arrow: matches metaKey+shiftKey', () => {
  assert.ok(KEYBOARD_PRESETS['cmd+shift+arrow']({ metaKey: true, ctrlKey: false, shiftKey: true, altKey: false }));
});

test('cmd+shift+arrow: matches ctrlKey+shiftKey', () => {
  assert.ok(KEYBOARD_PRESETS['cmd+shift+arrow']({ metaKey: false, ctrlKey: true, shiftKey: true, altKey: false }));
});

test('cmd+shift+arrow: rejects metaKey without shift', () => {
  assert.ok(!KEYBOARD_PRESETS['cmd+shift+arrow']({ metaKey: true, ctrlKey: false, shiftKey: false, altKey: false }));
});

test('cmd+shift+arrow: rejects altKey alongside meta+shift', () => {
  assert.ok(!KEYBOARD_PRESETS['cmd+shift+arrow']({ metaKey: true, ctrlKey: false, shiftKey: true, altKey: true }));
});

// ── getPresetHintText ─────────────────────────────────────────────────────────

test('getPresetHintText cmd+arrow Mac shows ⌘', () => {
  const text = getPresetHintText('cmd+arrow', 'MacIntel');
  assert.ok(text.includes('⌘'), 'should contain ⌘');
});

test('getPresetHintText cmd+arrow Win shows Ctrl', () => {
  const text = getPresetHintText('cmd+arrow', 'Win32');
  assert.ok(text.includes('Ctrl'), 'should contain Ctrl');
});

test('getPresetHintText alt+arrow Mac shows ⌥', () => {
  const text = getPresetHintText('alt+arrow', 'MacIntel');
  assert.ok(text.includes('⌥'), 'should contain ⌥');
});

test('getPresetHintText alt+arrow Win shows Alt+', () => {
  const text = getPresetHintText('alt+arrow', 'Win32');
  assert.ok(text.includes('Alt+'), 'should contain Alt+');
});

test('getPresetHintText cmd+shift+arrow Mac shows ⌘⇧', () => {
  const text = getPresetHintText('cmd+shift+arrow', 'MacIntel');
  assert.ok(text.includes('⌘⇧'), 'should contain ⌘⇧');
});

test('getPresetHintText cmd+shift+arrow Win shows Ctrl+Shift+', () => {
  const text = getPresetHintText('cmd+shift+arrow', 'Win32');
  assert.ok(text.includes('Ctrl+Shift+'), 'should contain Ctrl+Shift+');
});

test('getPresetHintText unknown preset falls back to cmd+arrow behaviour', () => {
  const text = getPresetHintText('unknown', 'MacIntel');
  assert.ok(text.includes('⌘'), 'unknown preset falls back to ⌘ on Mac');
});
