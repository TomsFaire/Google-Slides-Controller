# Configurable Keyboard Shortcuts Design

**Date:** 2026-05-14
**Branch:** `feature/configurable-keyboard-shortcuts`
**Status:** Approved for implementation

## Context

Web remote keyboard shortcuts (added in v2.3.2) are hardcoded to `Cmd/Ctrl+Arrow`. This conflicts with browser back/forward navigation in some browsers. Operators need to configure a safer preset before users connect, and optionally force-enable shortcuts by default so users don't have to tap the toggle.

---

## Feature Overview

Two new preferences stored in `preferences.json`:

- **`keyboardShortcutPreset`** — which key combo is active (string enum, default `"cmd+arrow"`)
- **`keyboardShortcutsDefaultEnabled`** — whether shortcuts start enabled when the web remote loads (boolean, default `false`)

Both are configurable from:
1. The **Electron desktop app** — Web Remote tab, new "Keyboard Shortcuts" panel
2. The **web remote Settings tab** — new "Keyboard Shortcuts" section

---

## Preset Options

| Preset key | Display name | Modifier check |
|---|---|---|
| `cmd+arrow` | `Cmd/Ctrl + Arrow` | `(e.metaKey \|\| e.ctrlKey) && !e.shiftKey && !e.altKey` |
| `alt+arrow` | `Alt/Option + Arrow` | `e.altKey && !e.metaKey && !e.ctrlKey && !e.shiftKey` |
| `cmd+shift+arrow` | `Cmd/Ctrl + Shift + Arrow` | `(e.metaKey \|\| e.ctrlKey) && e.shiftKey && !e.altKey` |

The arrow direction mapping stays the same across all presets:
- `ArrowLeft` → previous slide
- `ArrowRight` → next slide
- `ArrowUp` → scroll notes up
- `ArrowDown` → scroll notes down

---

## Architecture

### Preferences

`loadPreferences()` / `savePreferences()` in `main.js` already handle arbitrary preference keys. No schema changes needed — just read/write the two new keys.

### Template injection

When Electron renders the web remote HTML template, two JS variables are injected alongside the existing `window.__GSO_WEB_UI_RESTRICTED__`:

```js
window.__GSO_KEYBOARD_PRESET__ = '${prefs.keyboardShortcutPreset || "cmd+arrow"}';
window.__GSO_KEYBOARD_DEFAULT_ENABLED__ = ${!!prefs.keyboardShortcutsDefaultEnabled};
```

These are read once on page load to initialise `currentKeyboardPreset` and determine whether to auto-enable shortcuts.

### Runtime variable

A new `let currentKeyboardPreset` variable (alongside the existing `keyboardShortcutsEnabled`) holds the active preset string. The keydown handler does a lookup against `KEYBOARD_PRESETS` rather than a hardcoded modifier check:

```js
const KEYBOARD_PRESETS = {
  'cmd+arrow':       e => (e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey,
  'alt+arrow':       e => e.altKey && !e.metaKey && !e.ctrlKey && !e.shiftKey,
  'cmd+shift+arrow': e => (e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey,
};
```

When the user saves settings from the web Settings tab, `currentKeyboardPreset` is updated immediately so the new combo takes effect without a reload.

### Hint text

`updateKeyboardHint()` generates the hint string from the active preset. Friendly names per preset:

| Preset | Mac hint | Win/Linux hint |
|---|---|---|
| `cmd+arrow` | `⌘← ⌘→ ⌘↑ ⌘↓` | `Ctrl← Ctrl→ Ctrl↑ Ctrl↓` |
| `alt+arrow` | `⌥← ⌥→ ⌥↑ ⌥↓` | `Alt← Alt→ Alt↑ Alt↓` |
| `cmd+shift+arrow` | `⌘⇧← ⌘⇧→ ⌘⇧↑ ⌘⇧↓` | `Ctrl+Shift+← →` etc. |

`getShortcutModifier()` already detects Mac vs other. A new `getPresetHintText(preset)` helper builds the four-shortcut string.

### Default-enabled behaviour

On page load, after all JS initialises:

```js
const shouldEnable = window.__GSO_KEYBOARD_DEFAULT_ENABLED__ ||
  localStorage.getItem('gsc_keyboard_shortcuts_enabled') === 'true';
```

The admin default (`__GSO_KEYBOARD_DEFAULT_ENABLED__`) takes precedence: if it is `true`, shortcuts are enabled regardless of localStorage. If it is `false`, localStorage is the tiebreaker (preserving per-session user preference).

---

## Electron Desktop — Web Remote Tab

New **"Keyboard Shortcuts"** panel added to `index.html` under the Web Remote section, below the Stagetimer panel:

```html
<div class="panel">
  <h2 class="panel-title">Keyboard Shortcuts</h2>
  <div class="form-group">
    <label for="keyboard-shortcut-preset">Shortcut preset</label>
    <select id="keyboard-shortcut-preset" class="select-input">
      <option value="cmd+arrow">Cmd/Ctrl + Arrow</option>
      <option value="alt+arrow">Alt/Option + Arrow</option>
      <option value="cmd+shift+arrow">Cmd/Ctrl + Shift + Arrow (safest)</option>
    </select>
  </div>
  <div class="form-group">
    <label class="checkbox-label">
      <input type="checkbox" id="keyboard-shortcuts-default-enabled">
      Enable for new connections by default
    </label>
  </div>
  <button class="btn" id="btn-save-keyboard-shortcuts">Save Keyboard Settings</button>
</div>
```

`renderer.js` bindings:
- On load: read `prefs.keyboardShortcutPreset` and `prefs.keyboardShortcutsDefaultEnabled`, set dropdown and checkbox values
- On save: call `window.electronAPI.savePreferences({ keyboardShortcutPreset, keyboardShortcutsDefaultEnabled })`

---

## Web UI Settings Tab

New **"Keyboard Shortcuts"** section in the `#tab-settings` div, after the Stagetimer section:

```html
<div class="controls-section">
  <h3>Keyboard Shortcuts</h3>
  <div class="preset-group">
    <label for="web-keyboard-preset">Shortcut preset</label>
    <select id="web-keyboard-preset" class="input-field">
      <option value="cmd+arrow">Cmd/Ctrl + Arrow</option>
      <option value="alt+arrow">Alt/Option + Arrow</option>
      <option value="cmd+shift+arrow">Cmd/Ctrl + Shift + Arrow (safest)</option>
    </select>
  </div>
  <label class="checkbox-label">
    <input type="checkbox" id="web-keyboard-default-enabled">
    Enable for new connections by default
  </label>
  <button type="button" class="btn" id="btn-save-keyboard-shortcuts">
    Save Keyboard Settings
  </button>
</div>
```

On Settings tab load, the section reads its initial values from `window.__GSO_KEYBOARD_PRESET__` and `window.__GSO_KEYBOARD_DEFAULT_ENABLED__` (already injected). On save, calls `POST /api/preferences` with `{ keyboardShortcutPreset, keyboardShortcutsDefaultEnabled }` and then sets `currentKeyboardPreset` in the page JS so the new combo is live immediately.

---

## Files to Modify

- **`main.js`** — only file for web UI and API changes
  - Template injection: add two `window.__GSO_KEYBOARD_*__` vars alongside `window.__GSO_WEB_UI_RESTRICTED__`
  - Web Settings tab HTML: new "Keyboard Shortcuts" section
  - Web JS: replace hardcoded modifier check with `KEYBOARD_PRESETS` lookup; add `currentKeyboardPreset` var; update `updateKeyboardHint()` → `getPresetHintText()`; update localStorage restore + default-enabled logic; add Settings tab save handler
- **`index.html`** — Electron desktop settings
  - New "Keyboard Shortcuts" panel in the Web Remote section
- **`renderer.js`** — Electron desktop settings logic
  - Load preset and default-enabled values on init
  - Save handler calling `savePreferences()`

---

## API

No new endpoints. `POST /api/preferences` already accepts arbitrary preference keys and passes them through to `savePreferences()`.

---

## Verification

1. Open Electron desktop app → Web Remote tab → confirm "Keyboard Shortcuts" panel shows with dropdown and checkbox
2. Change preset to "Cmd/Ctrl + Shift + Arrow", save
3. Open web remote → Settings tab → confirm the dropdown shows "Cmd/Ctrl + Shift + Arrow"
4. In web remote Remote tab, enable keyboard shortcuts — confirm hint text shows `⌘⇧←` etc.
5. Press `Cmd+Shift+→` — presentation advances. Press old combo `Cmd+→` — nothing happens.
6. Check preset change to "Alt/Option + Arrow", save from web Settings tab — confirm keydown handler switches combo immediately (no reload)
7. Enable "Enable for new connections by default", save — reload web remote, confirm shortcuts are active on load without tapping the toggle
8. Disable "Enable for new connections by default", save — reload, confirm shortcuts start off
