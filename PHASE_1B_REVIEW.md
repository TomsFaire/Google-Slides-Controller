# Phase 1B - Codebase Review & Implementation Guide

**Date:** 2026-03-15
**Phase:** 1B (Layout Fix & Backup Controls)
**Tasks:** 2 (Speaker notes column width) + 3 (Backup controls toggle)

---

## Task 2: Fix Speaker Notes Preview Column Width

### Current State
- **Function:** `setSpeakerNotesFullscreen()` at main.js:427 — Sets notes window to fullscreen
- **Window setup:** `getSpeakerNotesWindowOptions()` at main.js:356 — Configures window size to full display bounds
- **Normalization:** `getNotesWindowNormalizeScript()` at main.js:375 — Injects JavaScript to fix text encoding
- **Issue:** Preview column expands to ~50% despite window being full-screen

### Architecture Notes
- Notes window is opened at full display bounds (line 357: `bounds.width`, `bounds.height`)
- Google Slides responsive layout: wide viewport = narrow preview column + wide notes
- JavaScript injection for DOM manipulation already works (used for text normalization)
- No existing CSS injection for preview column styling

### Implementation Approach
1. Create CSS injection function alongside `getNotesWindowNormalizeScript()`
2. Inject CSS into notes window that constrains preview column to 28% max-width
3. Apply CSS immediately on window load (in did-finish-load handler)
4. Target selector: Look for `.preview-column-selector` or similar in Google Slides presenter view DOM

### Key Code Locations
- Line 375: `getNotesWindowNormalizeScript()` — Existing JavaScript injection pattern
- Line 1188: `did-finish-load` → `setSpeakerNotesFullscreen()` — Window ready callback
- Line 406: `notesWindow.webContents.executeJavaScript()` — Method to run code in notes window

### Expected Changes
- Add `getCssForNotesWindow()` function (new, ~20 lines)
- Modify window load handler to inject CSS (1-2 lines)
- No API changes needed
- No Companion module changes needed

---

## Task 3: Enable/Disable Backup Controls Toggle

### Current State - Broadcasting Logic
- **Function:** `sendToBackups()` at main.js:985 — Broadcasts commands to backup machines
- **Status check:** `checkBackupStatus()` at main.js:1038 — Polls backup health
- **Preferences:** `loadPreferences()` at main.js:762 — Stores user config
- **No gate:** Currently broadcasts whenever `primaryBackupMode === 'primary'` (line 987)

### State Variables Needed
- New module-level variable: `let backupControlsEnabled = true;` (add near line 2570 with other state)
- Gate broadcasts: Add `if (!backupControlsEnabled) return;` inside `sendToBackups()` (line 988)

### API Endpoints Needed
1. **Expose in `/api/status`** (line 2556-2573)
   - Add field: `backupControlsEnabled: backupControlsEnabled,`

2. **New endpoint: `POST /api/set-backup-controls`**
   - Accept body: `{ "enabled": true|false }`
   - Set `backupControlsEnabled = data.enabled`
   - Return: `{ "success": true, "backupControlsEnabled": backupControlsEnabled }`
   - Error handling: Return 400 if missing required field

### Companion Module Changes
- **File:** `companion-module-gslide-opener/main.js`
- **Lines to modify:**
  - Line 231-296: `setupVariables()` — Add `backup_controls_enabled` variable definition
  - Line 303-402: `setupFeedbacks()` — Add feedback for backup controls enabled state
  - Line 415-493: `updateState()` — Poll and set `backup_controls_enabled` variable

- **File:** `companion-module-gslide-opener/actions.js`
  - Add new action: "Set Backup Controls" or "Enable/Disable Backup Controls"
  - Action should POST to `/api/set-backup-controls` with `{ "enabled": true|false }`

### Key Code Locations
- Line 985-1035: `sendToBackups()` function — Where to add gate
- Line 2555-2681: `/api/status` endpoint — Where to expose toggle
- Line 2520-2689: API status handler — Where to add new `/api/set-backup-controls` endpoint

### Implementation Approach
1. Add module-level `backupControlsEnabled = true;` (line ~2570)
2. Gate `sendToBackups()` with check at line 988
3. Add `/api/set-backup-controls` endpoint handler
4. Expose in `/api/status` response
5. Update Companion module: variable definition + variable polling + new action
6. Update CHANGELOG.md in both repos

### Expected Changes
- **main.js:** ~30 lines (state var + gate + endpoint handler + status field)
- **Companion main.js:** ~20 lines (variable definition + polling update)
- **Companion actions.js:** ~25 lines (new action definition)
- **CHANGELOG.md:** 2 entries (one per repo)

---

## Testing Checklist - Phase 1B

### Task 2 (Speaker Notes Width)
- [ ] Open presentation, open speaker notes window
- [ ] Verify preview column is constrained to ≤30% width
- [ ] Verify column does not expand as preview images load
- [ ] Test on multiple display sizes if possible
- [ ] Verify notes content is still readable

### Task 3 (Backup Controls)
- [ ] GET `/api/status` includes `backupControlsEnabled: true` (default)
- [ ] POST `/api/set-backup-controls` with `{ "enabled": false }` succeeds
- [ ] GET `/api/status` now shows `backupControlsEnabled: false`
- [ ] With toggle OFF: Commands sent to primary don't forward to backups
- [ ] With toggle ON: Commands broadcast normally to backups
- [ ] Companion module shows new variable `backup_controls_enabled`
- [ ] Companion action triggers API endpoint correctly
- [ ] Verify feedback (button color) changes based on toggle state

---

## Dependencies & Order
- Task 2 and Task 3 are **independent** — can be done in parallel
- Task 2: ~30 mins
- Task 3: ~45 mins
- Both should complete in single session with Sonnet

---

## Files to Modify Summary

```
main.js
├── ~30 lines: Add backupControlsEnabled state, gate in sendToBackups, add /api/set-backup-controls endpoint, expose in /api/status

companion-module-gslide-opener/main.js
├── ~20 lines: Variable definition + polling setup

companion-module-gslide-opener/actions.js
├── ~25 lines: New action definition

companion-module-gslide-opener/CHANGELOG.md
├── Entry for backup controls feature

CHANGELOG.md (main repo)
├── Entry for speaker notes width fix
└── Entry for backup controls feature
```

---

## Ready for Sonnet 4.6

This review document provides all context needed to execute Phase 1B with Sonnet 4.6.
