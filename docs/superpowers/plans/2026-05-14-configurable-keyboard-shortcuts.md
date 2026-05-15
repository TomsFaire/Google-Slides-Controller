# Configurable Keyboard Shortcuts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let operators configure the keyboard shortcut preset (`cmd+arrow`, `alt+arrow`, or `cmd+shift+arrow`) and whether shortcuts start enabled by default — from both the Electron desktop app's Web Remote tab and the web UI's Settings tab.

**Architecture:** All changes are confined to three files. `main.js` gets: (1) two new `window.__GSO_KEYBOARD_*__` vars injected into the HTML template, (2) a `KEYBOARD_PRESETS` map + `currentKeyboardPreset` runtime variable replacing the hardcoded modifier check, (3) a `getPresetHintText()` helper replacing the hardcoded hint string, (4) a new "Keyboard Shortcuts" section in the web Settings tab HTML + its init and save handler. `index.html` gets a new "Keyboard Shortcuts" panel in the Web Remote section. `renderer.js` gets DOM bindings to load and save the two new preferences via `window.electronAPI.savePreferences`.

**Tech Stack:** Vanilla JS, inline CSS template strings in `main.js`, Node.js built-in test runner (`node --test`)

---

## File Structure

- **Modify:** `main.js` — template injection, `KEYBOARD_PRESETS`, keydown refactor, hint helper, web Settings HTML + JS
- **Modify:** `index.html` — new Keyboard Shortcuts panel in Web Remote section
- **Modify:** `renderer.js` — DOM declarations, prefs restore, save function, button listener
- **Modify:** `tests/keyboard-shortcuts.test.js` — extend with KEYBOARD_PRESETS and getPresetHintText tests

---

## Task 1: Create feature branch

**Files:** none (git only)

- [ ] **Step 1: Create and check out the branch**

```bash
cd /Users/tom/Documents/gslide-opener
git checkout main
git checkout -b feature/configurable-keyboard-shortcuts
```

Expected: prompt shows `feature/configurable-keyboard-shortcuts`

- [ ] **Step 2: Verify clean state**

```bash
git status
```

Expected: `nothing to commit, working tree clean`

---

## Task 2: Add template injection for keyboard preferences

**Files:**
- Modify: `main.js` — line 9700 area

The web HTML template is a Node.js template literal. `prefs` (the loaded preferences object) and `webUiRestrictedTunnelClient` (a boolean) are both in scope at line 9700. We add two new injected globals directly after the existing `__GSO_WEB_UI_RESTRICTED__` line.

- [ ] **Step 1: Add the two injected globals**

Find this exact line in `main.js`:
```
    window.__GSO_WEB_UI_RESTRICTED__ = ${webUiRestrictedTunnelClient ? 'true' : 'false'};
```

Replace it with:
```
    window.__GSO_WEB_UI_RESTRICTED__ = ${webUiRestrictedTunnelClient ? 'true' : 'false'};
    window.__GSO_KEYBOARD_PRESET__ = '${prefs.keyboardShortcutPreset || "cmd+arrow"}';
    window.__GSO_KEYBOARD_DEFAULT_ENABLED__ = ${!!prefs.keyboardShortcutsDefaultEnabled};
```

- [ ] **Step 2: Commit**

```bash
git add main.js
git commit -m "feat: inject keyboard preset and default-enabled into web template"
```

---

## Task 3: Add KEYBOARD_PRESETS map and currentKeyboardPreset variable

**Files:**
- Modify: `main.js` — lines 9842-9873 area (keyboard shortcut JS state variables and helpers)

- [ ] **Step 1: Add KEYBOARD_PRESETS and currentKeyboardPreset after keyboardShortcutsEnabled**

Find this exact line in `main.js`:
```
    let keyboardShortcutsEnabled = false;
```

Replace it with:
```
    let keyboardShortcutsEnabled = false;
    const KEYBOARD_PRESETS = {
      'cmd+arrow':       e => (e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey,
      'alt+arrow':       e => e.altKey && !e.metaKey && !e.ctrlKey && !e.shiftKey,
      'cmd+shift+arrow': e => (e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey,
    };
    let currentKeyboardPreset = (window.__GSO_KEYBOARD_PRESET__ && KEYBOARD_PRESETS[window.__GSO_KEYBOARD_PRESET__])
      ? window.__GSO_KEYBOARD_PRESET__ : 'cmd+arrow';
```

- [ ] **Step 2: Add getPresetHintText helper and update updateKeyboardHint**

Find this exact block in `main.js`:
```
    function updateKeyboardHint() {
      const hint = document.getElementById('keyboard-shortcuts-hint');
      const keysEl = document.getElementById('keyboard-hint-keys');
      const mod = getShortcutModifier();
      if (!hint || !keysEl) return;
      if (keyboardShortcutsEnabled) {
        keysEl.textContent = mod + '← Prev slide · ' + mod + '→ Next slide · ' + mod + '↑ Notes up · ' + mod + '↓ Notes down';
        hint.classList.add('visible');
      } else {
        hint.classList.remove('visible');
      }
    }
```

Replace it with:
```
    function getPresetHintText(preset) {
      const mod = getShortcutModifier();
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

    function updateKeyboardHint() {
      const hint = document.getElementById('keyboard-shortcuts-hint');
      const keysEl = document.getElementById('keyboard-hint-keys');
      if (!hint || !keysEl) return;
      if (keyboardShortcutsEnabled) {
        keysEl.textContent = getPresetHintText(currentKeyboardPreset);
        hint.classList.add('visible');
      } else {
        hint.classList.remove('visible');
      }
    }
```

- [ ] **Step 3: Update localStorage restore IIFE to honour admin default-enabled**

Find this exact block in `main.js`:
```
    // Restore keyboard shortcut toggle from localStorage
    (function() {
      const stored = localStorage.getItem('gsc_keyboard_shortcuts_enabled');
      if (stored === 'true') {
        keyboardShortcutsEnabled = true;
        const btn = document.getElementById('keyboard-toggle-btn');
        if (btn) btn.classList.add('active');
        updateKeyboardHint();
      }
    })();
```

Replace it with:
```
    // Restore keyboard shortcut toggle from admin default or localStorage
    (function() {
      const adminDefault = window.__GSO_KEYBOARD_DEFAULT_ENABLED__ === true;
      const stored = localStorage.getItem('gsc_keyboard_shortcuts_enabled');
      const shouldEnable = adminDefault || stored === 'true';
      if (shouldEnable) {
        keyboardShortcutsEnabled = true;
        const btn = document.getElementById('keyboard-toggle-btn');
        if (btn) btn.classList.add('active');
        updateKeyboardHint();
      }
    })();
```

- [ ] **Step 4: Replace hardcoded modifier check in keydown listener with KEYBOARD_PRESETS lookup**

Find this exact block in `main.js`:
```
    document.addEventListener('keydown', function(e) {
      if (!keyboardShortcutsEnabled) return;
      const tag = document.activeElement ? document.activeElement.tagName : '';
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || (document.activeElement && document.activeElement.isContentEditable)) return;
      const isMod = e.metaKey || e.ctrlKey;
      if (!isMod) return;
      if (e.key === 'ArrowRight') {
```

Replace it with:
```
    document.addEventListener('keydown', function(e) {
      if (!keyboardShortcutsEnabled) return;
      const tag = document.activeElement ? document.activeElement.tagName : '';
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || (document.activeElement && document.activeElement.isContentEditable)) return;
      const checker = KEYBOARD_PRESETS[currentKeyboardPreset] || KEYBOARD_PRESETS['cmd+arrow'];
      if (!checker(e)) return;
      if (e.key === 'ArrowRight') {
```

- [ ] **Step 5: Commit**

```bash
git add main.js
git commit -m "feat: refactor keydown handler to use KEYBOARD_PRESETS map"
```

---

## Task 4: Add web Settings tab Keyboard Shortcuts section

**Files:**
- Modify: `main.js` — web Settings HTML (after Logging section ~line 9649) and web JS (after Logging save handler ~line 11953, and in prefs-load block ~line 11698)

### 4a — HTML

- [ ] **Step 1: Add Keyboard Shortcuts HTML section after the Logging section**

Find this exact line in `main.js`:
```
        <button type="button" class="btn" id="btn-save-logging" style="margin-top: 12px;">Save Logging Settings</button>
      </div>
      
      ${webUiDebugConsoleEnabled ? `
```

Replace it with:
```
        <button type="button" class="btn" id="btn-save-logging" style="margin-top: 12px;">Save Logging Settings</button>
      </div>

      <!-- Keyboard Shortcuts Section -->
      <div class="controls-section">
        <h3>Keyboard Shortcuts</h3>
        <div class="info" style="margin-bottom: 10px;">
          Configure the keyboard combo used when shortcuts are enabled. Changes take effect immediately without a page reload.
        </div>
        <div class="preset-group">
          <label for="web-keyboard-preset">Shortcut preset</label>
          <select id="web-keyboard-preset" class="input-field">
            <option value="cmd+arrow">Cmd/Ctrl + Arrow</option>
            <option value="alt+arrow">Alt/Option + Arrow</option>
            <option value="cmd+shift+arrow">Cmd/Ctrl + Shift + Arrow (safest)</option>
          </select>
        </div>
        <div style="display: flex; align-items: center; gap: 10px; margin-top: 12px;">
          <input type="checkbox" id="web-keyboard-default-enabled" style="width: auto;" />
          <label for="web-keyboard-default-enabled" style="margin: 0; font-weight: normal;">Enable for new connections by default</label>
        </div>
        <button type="button" class="btn" id="btn-save-keyboard-shortcuts" style="margin-top: 12px;">Save Keyboard Settings</button>
      </div>
      
      ${webUiDebugConsoleEnabled ? `
```

### 4b — Init from prefs

- [ ] **Step 2: Initialise the new controls when the Settings tab loads its prefs**

Find this exact block in `main.js` (the verbose logging init, inside the Settings tab init block):
```
        // Set logging preferences
        const verboseEl = document.getElementById('web-verbose-logging');
        if (verboseEl) {
          verboseEl.checked = prefs.verboseLogging === true;
        }
        
        // Set up primary/backup mode change handlers
```

Replace it with:
```
        // Set logging preferences
        const verboseEl = document.getElementById('web-verbose-logging');
        if (verboseEl) {
          verboseEl.checked = prefs.verboseLogging === true;
        }

        // Set keyboard shortcut preferences
        const webKeyboardPresetEl = document.getElementById('web-keyboard-preset');
        if (webKeyboardPresetEl) {
          const validPresets = ['cmd+arrow', 'alt+arrow', 'cmd+shift+arrow'];
          webKeyboardPresetEl.value = validPresets.includes(prefs.keyboardShortcutPreset) ? prefs.keyboardShortcutPreset : 'cmd+arrow';
        }
        const webKeyboardDefaultEl = document.getElementById('web-keyboard-default-enabled');
        if (webKeyboardDefaultEl) {
          webKeyboardDefaultEl.checked = prefs.keyboardShortcutsDefaultEnabled === true;
        }

        // Set up primary/backup mode change handlers
```

### 4c — Save handler

- [ ] **Step 3: Add the save handler after the Logging save handler**

Find this exact block in `main.js`:
```
    // Backup status polling
    let webBackupStatusInterval = null;
```

Insert before it:
```
    // Save keyboard shortcut settings
    const saveKeyboardShortcutsBtn = document.getElementById('btn-save-keyboard-shortcuts');
    if (saveKeyboardShortcutsBtn) {
      saveKeyboardShortcutsBtn.addEventListener('click', async () => {
        try {
          const presetEl = document.getElementById('web-keyboard-preset');
          const defaultEl = document.getElementById('web-keyboard-default-enabled');
          const validPresets = ['cmd+arrow', 'alt+arrow', 'cmd+shift+arrow'];
          const preset = presetEl && validPresets.includes(presetEl.value) ? presetEl.value : 'cmd+arrow';
          const defaultEnabled = defaultEl ? defaultEl.checked : false;
          const res = await fetch(API_BASE + '/api/preferences', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ keyboardShortcutPreset: preset, keyboardShortcutsDefaultEnabled: defaultEnabled })
          });
          const result = await res.json();
          if (result.success) {
            currentKeyboardPreset = preset;
            updateKeyboardHint();
            showStatus('Keyboard settings saved', false);
          } else {
            showStatus('Failed to save keyboard settings: ' + (result.error || 'Unknown error'), true);
          }
        } catch (error) {
          showStatus('Failed to save keyboard settings: ' + error.message, true);
        }
      });
    }

    // Backup status polling
    let webBackupStatusInterval = null;
```

- [ ] **Step 4: Commit**

```bash
git add main.js
git commit -m "feat: add Keyboard Shortcuts section to web Settings tab"
```

---

## Task 5: Add Electron desktop Keyboard Shortcuts panel

**Files:**
- Modify: `index.html` — new panel in Web Remote section (before closing `</section>` at line ~434)

- [ ] **Step 1: Add the Keyboard Shortcuts panel**

Find this exact line in `index.html`:
```
            <button id="save-stagetimer-btn" class="btn btn-primary">Save Stagetimer Settings</button>
            <button id="load-stagetimer-btn" class="btn btn-secondary" style="margin-top: 8px;">Load Current Settings</button>
          </div>
        </section>
```

Replace it with:
```
            <button id="save-stagetimer-btn" class="btn btn-primary">Save Stagetimer Settings</button>
            <button id="load-stagetimer-btn" class="btn btn-secondary" style="margin-top: 8px;">Load Current Settings</button>
          </div>

          <!-- Keyboard Shortcuts panel -->
          <div class="panel">
            <h2 class="panel-title">Keyboard Shortcuts</h2>
            <p class="card-description">Configure the default keyboard shortcut preset for the web remote. Users can still toggle shortcuts on/off per session.</p>

            <div class="form-group">
              <label for="keyboard-shortcut-preset">Shortcut preset</label>
              <select id="keyboard-shortcut-preset" class="select-input">
                <option value="cmd+arrow">Cmd/Ctrl + Arrow</option>
                <option value="alt+arrow">Alt/Option + Arrow</option>
                <option value="cmd+shift+arrow">Cmd/Ctrl + Shift + Arrow (safest)</option>
              </select>
            </div>

            <div class="form-group">
              <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                <input type="checkbox" id="keyboard-shortcuts-default-enabled" style="width: auto;" />
                <span>Enable for new connections by default</span>
              </label>
            </div>

            <button id="save-keyboard-shortcuts-btn" class="btn btn-primary">Save Keyboard Settings</button>
          </div>
        </section>
```

- [ ] **Step 2: Commit**

```bash
git add index.html
git commit -m "feat: add Keyboard Shortcuts panel to Electron desktop Web Remote tab"
```

---

## Task 6: Wire up renderer.js

**Files:**
- Modify: `renderer.js` — DOM declarations, prefs restore, save function, button listener

### 6a — DOM element declarations

- [ ] **Step 1: Add DOM element constants**

Find this exact line in `renderer.js` (line ~43):
```
const saveWebUiAppearanceBtn = document.getElementById('save-web-ui-appearance-btn');
```

Insert immediately after it:
```
const keyboardShortcutPresetSelect = document.getElementById('keyboard-shortcut-preset');
const keyboardShortcutsDefaultEnabledCheckbox = document.getElementById('keyboard-shortcuts-default-enabled');
const saveKeyboardShortcutsBtn = document.getElementById('save-keyboard-shortcuts-btn');
```

### 6b — Prefs restore

- [ ] **Step 2: Restore keyboard prefs from preferences object**

Find this exact block in `renderer.js`:
```
    // Restore Web UI appearance (theme + logo + custom CSS path)
    if (webUiThemeSelect) {
```

Insert before it:
```
    // Restore keyboard shortcut preferences
    if (keyboardShortcutPresetSelect) {
      const validPresets = ['cmd+arrow', 'alt+arrow', 'cmd+shift+arrow'];
      const preset = preferences.keyboardShortcutPreset;
      keyboardShortcutPresetSelect.value = validPresets.includes(preset) ? preset : 'cmd+arrow';
    }
    if (keyboardShortcutsDefaultEnabledCheckbox) {
      keyboardShortcutsDefaultEnabledCheckbox.checked = preferences.keyboardShortcutsDefaultEnabled === true;
    }

```

### 6c — Save function

- [ ] **Step 3: Add saveKeyboardShortcutSettings function**

Find this exact line in `renderer.js`:
```
// Save primary/backup preferences
async function savePrimaryBackupPreferences() {
```

Insert before it:
```
async function saveKeyboardShortcutSettings() {
  try {
    const validPresets = ['cmd+arrow', 'alt+arrow', 'cmd+shift+arrow'];
    const preset = keyboardShortcutPresetSelect ? keyboardShortcutPresetSelect.value : 'cmd+arrow';
    const defaultEnabled = keyboardShortcutsDefaultEnabledCheckbox ? keyboardShortcutsDefaultEnabledCheckbox.checked : false;
    await window.electronAPI.savePreferences({
      keyboardShortcutPreset: validPresets.includes(preset) ? preset : 'cmd+arrow',
      keyboardShortcutsDefaultEnabled: defaultEnabled
    });
    showStatus('Keyboard shortcut settings saved.', 'info');
  } catch (error) {
    console.error('Failed to save keyboard shortcut settings:', error);
    showStatus('Failed to save keyboard shortcut settings', 'error');
  }
}

```

### 6d — Button listener

- [ ] **Step 4: Wire the button**

Find this exact block in `renderer.js`:
```
    if (saveWebUiAppearanceBtn) {
      saveWebUiAppearanceBtn.addEventListener('click', saveWebUiAppearance);
    }
```

Insert immediately after it:
```
    if (saveKeyboardShortcutsBtn) {
      saveKeyboardShortcutsBtn.addEventListener('click', saveKeyboardShortcutSettings);
    }
```

- [ ] **Step 5: Commit**

```bash
git add renderer.js
git commit -m "feat: add keyboard shortcut settings bindings to renderer.js"
```

---

## Task 7: Extend tests

**Files:**
- Modify: `tests/keyboard-shortcuts.test.js`

- [ ] **Step 1: Replace the test file with the extended version**

Read the current file first to confirm it exists, then overwrite it with the content below.

Create `tests/keyboard-shortcuts.test.js` with this content:
```js
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
```

- [ ] **Step 2: Run the tests**

```bash
export NVM_DIR="$HOME/.nvm" && source "$NVM_DIR/nvm.sh"
node --test tests/keyboard-shortcuts.test.js
```

Expected: 27 passing tests, 0 failures.

- [ ] **Step 3: Commit**

```bash
git add tests/keyboard-shortcuts.test.js
git commit -m "test: extend keyboard shortcuts tests for KEYBOARD_PRESETS and getPresetHintText"
```

---

## Task 8: Bump version and update CHANGELOG

**Files:**
- Modify: `package.json` — version + buildNumber
- Modify: `CHANGELOG.md` — new v2.3.3 entry

- [ ] **Step 1: Bump version to 2.3.3, build 79**

In `package.json`, find:
```json
  "version": "2.3.2",
  "buildNumber": "78",
```

Replace with:
```json
  "version": "2.3.3",
  "buildNumber": "79",
```

- [ ] **Step 2: Add CHANGELOG entry**

In `CHANGELOG.md`, find the `## [2.3.2]` heading and insert the new entry above it:

```markdown
## [2.3.3] - 2026-05-14

### Added
- **Configurable keyboard shortcut preset** — Operators can now choose between three presets from the Electron desktop app's Web Remote tab or the web remote's Settings tab:
  - `Cmd/Ctrl + Arrow` (original, default)
  - `Alt/Option + Arrow`
  - `Cmd/Ctrl + Shift + Arrow` (safest — avoids browser back/forward conflict)
- **Default-enabled toggle** — Admins can pre-enable keyboard shortcuts for all new connections without requiring users to tap the toggle.
- Both settings persist in `preferences.json` and are applied on web remote load via injected template globals.
- Preset changes in the web Settings tab take effect immediately without a page reload.

```

- [ ] **Step 3: Commit**

```bash
git add package.json CHANGELOG.md
git commit -m "chore: bump to v2.3.3 (build 79), add CHANGELOG entry for configurable keyboard shortcuts"
```

---

## Task 9: Manual verification

- [ ] **Step 1: Start the app**

```bash
export NVM_DIR="$HOME/.nvm" && source "$NVM_DIR/nvm.sh"
cd /Users/tom/Documents/gslide-opener
yarn start
```

- [ ] **Step 2: Verify Electron desktop panel**

Open the desktop app → Web Remote tab. Confirm:
- "Keyboard Shortcuts" panel appears below Stagetimer with a preset dropdown and a checkbox.

- [ ] **Step 3: Change preset and save from Electron**

Select "Cmd/Ctrl + Shift + Arrow (safest)", click "Save Keyboard Settings". No error status.

- [ ] **Step 4: Verify web Settings tab reflects saved preset**

Open web remote → Settings tab. Confirm the "Shortcut preset" dropdown shows "Cmd/Ctrl + Shift + Arrow".

- [ ] **Step 5: Verify hint text updates to new preset**

In web remote Remote tab, enable keyboard shortcuts via the keyboard icon button. Confirm hint text shows `⌘⇧←` etc. (macOS) or `Ctrl+Shift+←` etc. (Windows/Linux).

- [ ] **Step 6: Verify new preset fires and old one does not**

Press `Cmd+Shift+→` — presentation advances. Press `Cmd+→` — nothing happens.

- [ ] **Step 7: Change preset from web Settings tab**

In web Settings tab, change preset to "Alt/Option + Arrow", click "Save Keyboard Settings". Confirm:
- Status shows "Keyboard settings saved"
- Hint text immediately updates to show `⌥←` etc. without a page reload

- [ ] **Step 8: Verify alt+arrow fires**

Press `Alt+→` (or `Option+→` on Mac) — presentation advances. Press `Cmd+→` — nothing happens.

- [ ] **Step 9: Test default-enabled**

Enable "Enable for new connections by default" in Electron desktop Web Remote tab, save. Reload the web remote. Confirm shortcuts are active on load (keyboard icon shows active, hint text visible) without tapping the toggle.

- [ ] **Step 10: Disable default-enabled**

Disable "Enable for new connections by default", save. Reload web remote. Confirm shortcuts start inactive.

---

## Task 10: Open pull request

- [ ] **Step 1: Push the branch**

```bash
git push -u origin feature/configurable-keyboard-shortcuts
```

- [ ] **Step 2: Open the PR**

```bash
gh pr create \
  --title "feat: configurable keyboard shortcut preset and default-enabled" \
  --body "$(cat <<'EOF'
## Summary
- Operators can now select the keyboard shortcut combo from three presets — from the Electron Web Remote tab or the web UI Settings tab
- Three presets: Cmd/Ctrl+Arrow (default), Alt/Option+Arrow, Cmd/Ctrl+Shift+Arrow (safest, avoids browser nav conflict)
- New \"Enable for new connections by default\" toggle lets admins pre-enable shortcuts before users connect
- Preset changes in the web Settings tab take effect immediately (no reload)
- Both prefs persisted in preferences.json and injected into the web page template on load

## Test plan
- [ ] Electron Web Remote tab shows new Keyboard Shortcuts panel with dropdown and checkbox
- [ ] Changing preset in Electron and saving is reflected in web Settings tab on next load
- [ ] Changing preset in web Settings tab takes effect immediately (hint text updates, new combo fires, old combo stops)
- [ ] Default-enabled: reload web remote, shortcuts auto-active when enabled; auto-inactive when disabled
- [ ] `node --test tests/keyboard-shortcuts.test.js` passes (27 tests)

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

---

## Spec reference

Full design doc: `docs/superpowers/specs/2026-05-14-configurable-keyboard-shortcuts-design.md`
