# Web Remote V2-C — Implementation Plan

**Target:** refresh the web UI served by the Electron main process (the remote page served by `GET /`) to the **V2-C "Stage-ready, notes-expanded"** design.

**Visual reference:** `./Web Remote UI.html` (in this same folder). Open it in a browser and look at the V2-C artboard — that is the ground truth for the layout, typography, and Stagetimer states described below.

**Supporting files in this folder:**
- `Web Remote UI.html` — design canvas with all variants, including V2-C
- `design-canvas.jsx` — canvas scaffolding (not used by the app; only for viewing the mock)
- `components/web-remote-variants.jsx` — V0/V1/V2/V3 artboard source
- `components/v2-explorations.jsx` — V2-A/B/C artboard source (V2-C is the target)

---

## 0. Guardrails

- **Do not break the REST API.** Endpoints (`/api/slide-prev`, `/api/slide-next`, `/api/notes-scroll-up`, `/api/notes-scroll-down`, `/api/get-stagetimer-status`, socket.io events, etc.) are consumed by the Companion module and Chrome extension. Do not rename them.
- **Do not change preferences schema.** `web-ui-theme`, `web-ui-logo-*`, `web-ui-custom-css-*`, `stagetimer-*`, `default-notes-zoom-steps` all stay. This is a CSS/HTML refactor inside the `theme-light` branch (or a new `theme-faire` option — see §1).
- **Do not change Settings / Controls tabs content.** Only re-style them to match. Only the **Remote** tab is restructured.
- **Preserve every DOM id.** The existing event wiring binds by id:
  - `stagetimer-container`, `stagetimer-label`, `stagetimer-time`, `stagetimer-status`, `stagetimer-messages`
  - `btn-scroll-notes-up`, `btn-scroll-notes-down`, `notes-zoom-in`, `notes-zoom-out`, `notes-zoom-controls`
  - `speaker-notes-content`, `speaker-notes-content-wrapper`, `speaker-notes-container`
  - `slide-preview-current-card`, `slide-preview-current-img`, `slide-preview-current-label`, `slide-preview-next-card`, `slide-preview-next-img`, `slide-preview-next-label`
  - `slide-previews-container`, `slide-previews-grid`
  - `remote-btn-prev`, `remote-btn-next`, `remote-controls`
  - `notes-toggle-btn`, `previews-toggle-btn`
  - All tabs: `tab-remote`, `tab-controls`, `tab-settings`, `.tab-btn[data-tab]`

  Moving an element in the DOM is fine. Renaming its id is not.

## 1. Theme strategy

Two options — pick one:

**Option A (recommended):** rewrite the `body.theme-light` block to the new Faire-flavored styling. Anyone with `web-ui-theme=light` gets the upgrade automatically. "Light (minimalist white)" in the select stays.

**Option B:** add a new theme option `faire` alongside the existing five. Adds a row to the `web-ui-theme` `<select>` in `index.html` (Electron preferences) and in the inline preferences HTML inside `main.js`. Default stays `light`; opt-in migration.

Plan below assumes **Option A**.

## 2. Where the code lives (use search, not line numbers)

The remote page is **not** a standalone file and **not** a named function. It is an inline template literal inside the `GET /` Express route in `main.js`. Do not chase line numbers — they drift. Search for these anchors instead:

| What you're editing | Search anchor (in `main.js`) |
|---|---|
| Remote-page route start | `GET / - Serve the web UI` |
| Base component styles (buttons, stagetimer base) | `.stagetimer-container {` (the big one with `linear-gradient(135deg, #667eea`) |
| Theme overrides | `body.theme-light { background: #f5f5f5` |
| Remote-tab HTML | `<div id="tab-remote" class="tab-content active">` |
| Stagetimer update (REST path) | `function updateStagetimerDisplay(` |
| Stagetimer update (socket.io path) | `function updateStagetimerDisplayFromState(` |
| Zoom handlers | `document.getElementById('notes-zoom-out').addEventListener` |
| Scroll-notes handlers | `document.getElementById('btn-scroll-notes-up').addEventListener` |

No changes needed in: `index.html`, `renderer.js`, `preload.js`, preferences, IPC.

## 3. Design tokens

Add these CSS custom properties at the top of the `<style>` block in the remote-page template so all new rules key off them:

```css
:root {
  --faire-font-sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  --faire-font-serif: 'Lora', Georgia, 'Times New Roman', serif;
  --faire-font-mono: ui-monospace, 'SF Mono', Menlo, Consolas, monospace;

  --faire-text: #333333;
  --faire-sub: #757575;
  --faire-muted: #b5a998;
  --faire-border: #dfe0e1;
  --faire-surface: #ffffff;
  --faire-warm: #fbf8f6;
  --faire-page: #fafaf8;

  /* Stagetimer tone tokens */
  --tmr-idle-bg: #fbf8f6;    --tmr-idle-bd: #dfe0e1;  --tmr-idle-fg: #757575;  --tmr-idle-clk: #333333;
  --tmr-run-bg:  #eef2ed;    --tmr-run-bd:  #c8d4c8;  --tmr-run-fg:  #49694c;  --tmr-run-clk:  #2d4a30;
  --tmr-warn-bg: #f6efdb;    --tmr-warn-bd: #d1b985;  --tmr-warn-fg: #907c3a;  --tmr-warn-clk: #5c4e1e;
  --tmr-crit-bg: #f5dcd6;    --tmr-crit-bd: #d9a79a;  --tmr-crit-fg: #921100;  --tmr-crit-clk: #6e1100;
  --tmr-over-bg: #3a1510;    --tmr-over-bd: #6e1100;  --tmr-over-fg: #ffd3c9;  --tmr-over-clk: #ffffff;

  --faire-radius: 4px;
}
```

## 4. Remote-tab DOM restructure (V2-C)

Rewrite the `<div id="tab-remote" class="tab-content active">` block so the order top→bottom is:

1. **Header row** (new wrapper `<div class="remote-header-compact">`)
   - Status dot + `machineName` in `--faire-text` 500
   - Slide counter "3 / 24" right-aligned, `--faire-sub`
   - `#notes-toggle-btn` and `#previews-toggle-btn` move here as icon-only 32×32 buttons (keep their ids and click handlers).
2. **`#stagetimer-container`** (stays always-mounted so no layout shift; CSS handles hidden state — see §5).
3. **Slide strip** (replaces `.slide-previews-grid` visuals but keeps all ids).
   - Same grid ids/structure but each card becomes `[thumbnail] [Label / slide title]` horizontally.
   - The thumb image keeps `#slide-preview-current-img` / `#slide-preview-next-img`.
   - Reuse `#slide-preview-current-label` / `#slide-preview-next-label` for the title text.
4. **Notes card** (`#speaker-notes-container`, becomes `flex: 1`)
   - Toolbar row: "Speaker notes · slide N" on the left; scroll cluster (↑/↓) and zoom cluster (−/px-readout/+) on the right.
   - Keep all four buttons (`#btn-scroll-notes-up`, `#btn-scroll-notes-down`, `#notes-zoom-out`, `#notes-zoom-in`) with their ids; replace text labels ("Scroll Up") with SVG icons + `aria-label`.
   - Add `<span id="notes-zoom-readout">18px</span>` between `#notes-zoom-out` and `#notes-zoom-in`; update it in the zoom handlers (see §6).
   - Wrapper body stays `#speaker-notes-content-wrapper > #speaker-notes-content` — only the typography changes (Lora 19/30).
5. **`#remote-controls`** — two big buttons, `Previous` secondary and `Next slide` primary, 72px tall.
6. **Bottom tab bar** (replaces the top `.tabs` row for `theme-light` only). The three `.tab-btn` buttons with their `data-tab` attributes are mirrored in a `<nav class="bottom-tabs">` at the bottom of the `.container`. See §8.

## 5. Stagetimer — state-driven colors

The existing code already sets one of `.stagetimer-container`, `.running`, `.warning` (remainingSeconds ≤ 60), `.critical` (remainingSeconds ≤ 15), `.error`, or `.disabled`.

### 5.1. Add the `overtime` state — both update functions

There are **two** stagetimer updaters and they compute `remainingMs` differently:

- `updateStagetimerDisplay(data, errorMessage)` — REST fallback path. Uses `data.remainingMs` (pre-computed server-side).
- `updateStagetimerDisplayFromState()` — live socket.io path (no parameters; uses global `stagetimerState`). Computes `remainingMs` **locally** from timer start/duration/elapsed (~20 lines before the className assignments). Use that local variable; do not copy-paste a `data.remainingMs` reference into this function.

In **both** functions, the overtime branch must come **before** the existing 15s / 60s branches, since `-5` is also `<= 15`. Pattern:

```js
// in updateStagetimerDisplay — data.remainingMs
if (data.remainingMs !== undefined) {
  const remainingSeconds = Math.floor(data.remainingMs / 1000);
  if (remainingSeconds < 0) {
    container.className = 'stagetimer-container overtime';
  } else if (remainingSeconds <= 15) {
    container.className = 'stagetimer-container critical';
  } else if (remainingSeconds <= 60) {
    container.className = 'stagetimer-container warning';
  }
}
```

```js
// in updateStagetimerDisplayFromState — local remainingMs (already computed above)
const remainingSeconds = Math.floor(remainingMs / 1000);
if (remainingSeconds < 0) {
  container.className = 'stagetimer-container overtime';
} else if (remainingSeconds <= 15) {
  container.className = 'stagetimer-container critical';
} else if (remainingSeconds <= 60) {
  container.className = 'stagetimer-container warning';
}
```

Keep the existing 15s / 60s thresholds unless product confirms otherwise.

### 5.2. Rewrite stagetimer CSS for `theme-light` (drop the gradient)

```css
body.theme-light .stagetimer-container {
  background: var(--tmr-idle-bg);
  border: 1px solid var(--tmr-idle-bd);
  color: var(--tmr-idle-fg);
  border-radius: var(--faire-radius);
  box-shadow: none;
  padding: 12px 14px;
  height: auto;        /* kill the 160px fixed height */
  text-align: left;
  display: flex;
  flex-direction: column;
  gap: 10px;
}
body.theme-light .stagetimer-container.running  { background: var(--tmr-run-bg);  border-color: var(--tmr-run-bd);  color: var(--tmr-run-fg); }
body.theme-light .stagetimer-container.warning  { background: var(--tmr-warn-bg); border-color: var(--tmr-warn-bd); color: var(--tmr-warn-fg); }
body.theme-light .stagetimer-container.critical { background: var(--tmr-crit-bg); border-color: var(--tmr-crit-bd); color: var(--tmr-crit-fg); }
body.theme-light .stagetimer-container.overtime { background: var(--tmr-over-bg); border-color: var(--tmr-over-bd); color: var(--tmr-over-fg); }
body.theme-light .stagetimer-container.error    { background: var(--tmr-crit-bg); border-color: var(--tmr-crit-bd); color: var(--tmr-crit-fg); }
body.theme-light .stagetimer-container.disabled { background: var(--tmr-idle-bg); border-color: var(--tmr-idle-bd); color: var(--faire-muted); }

body.theme-light .stagetimer-time {
  font-family: var(--faire-font-mono);
  font-size: 28px; font-weight: 500; letter-spacing: 1px;
  font-variant-numeric: tabular-nums;
  color: currentColor;
}
body.theme-light .stagetimer-label {
  font-size: 10.5px; letter-spacing: 0.7px; text-transform: uppercase; font-weight: 500;
}
body.theme-light .stagetimer-status { font-size: 11.5px; }

/* Messages — drop gradient scrim, reuse tone-matched border-top */
body.theme-light .stagetimer-messages {
  position: static;     /* no absolute-positioning in the new layout */
  border-top: 1px solid currentColor;
  background: transparent;
  backdrop-filter: none;
  padding: 10px 0 0;
  max-height: none;
}
body.theme-light .stagetimer-message {
  background: transparent; color: currentColor;
  padding: 0; font-size: 12.5px; line-height: 17px;
}
```

### 5.3. Rework markup inside `#stagetimer-container`

The existing layout has the clock centered and messages absolutely positioned (explains the 160px fixed height). In V2-C, header row + messages flow vertically, card grows if a message appears — no layout shift.

```html
<div class="stagetimer-container disabled" id="stagetimer-container" style="display: none;">
  <div class="stagetimer-row">
    <span class="stagetimer-dot"></span>
    <div class="stagetimer-label" id="stagetimer-label">Stage timer</div>
    <div class="stagetimer-time" id="stagetimer-time">--:--</div>
  </div>
  <div class="stagetimer-status" id="stagetimer-status"></div>
  <div class="stagetimer-messages" id="stagetimer-messages"></div>
</div>
```

Existing JS writes `textContent` to all three ids — no JS refactor beyond §5.1. `stagetimer-dot` is purely decorative.

## 6. Notes zoom readout

Existing handlers set `speakerNotesContent.style.fontSize`. After each change, also:

```js
const readout = document.getElementById('notes-zoom-readout');
if (readout) {
  const px = parseInt(getComputedStyle(speakerNotesContent).fontSize, 10);
  readout.textContent = px + 'px';
}
```

Call it once on init (after restoring persisted zoom level from localStorage).

## 7. CSS for the Remote tab (theme-light branch)

New block (replaces current `body.theme-light .*` rules for Remote-tab elements):

```css
body.theme-light {
  background: var(--faire-page);
  padding-top: 0;           /* kill 25vh padding — we want full-height flex */
  font-family: var(--faire-font-sans);
  color: var(--faire-text);
}
body.theme-light .container {
  max-width: min(440px, 100vw);
  background: var(--faire-surface);
  border: 1px solid var(--faire-border);
  border-radius: 0; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
  padding: 0;              /* inner sections own their padding */
  display: flex; flex-direction: column;
  min-height: 100vh;
}

/* Header */
body.theme-light .remote-header-compact {
  display: flex; align-items: center; gap: 10px;
  padding: 14px 18px 10px; font-size: 11.5px;
}
body.theme-light .remote-header-compact h1 { font-size: 14px; font-weight: 500; margin: 0; }
body.theme-light .remote-header-compact .slide-counter {
  margin-left: auto; color: var(--faire-sub); font-variant-numeric: tabular-nums;
}
body.theme-light .notes-toggle-btn,
body.theme-light .preview-toggle-btn {
  width: 32px; height: 32px; padding: 0; border-radius: var(--faire-radius);
  border: 1px solid var(--faire-border); background: var(--faire-surface); color: var(--faire-sub);
}

/* Hide top .tabs on light theme — tabs live at the bottom instead */
body.theme-light > .container > .tabs { display: none; }

/* Slide strip */
body.theme-light .slide-previews-grid {
  display: grid; grid-template-columns: 1fr 1fr; gap: 10px;
  padding: 0 18px 12px; background: transparent; border: none;
}
body.theme-light .slide-preview-card {
  display: flex; gap: 8px; align-items: center;
  padding: 0; background: transparent; border: none;
}
body.theme-light .slide-preview-img { width: 96px; height: 54px; border-radius: 2px; border: 1px solid var(--faire-border); object-fit: cover; }
body.theme-light .slide-preview-card:last-child .slide-preview-img { width: 72px; height: 40px; }
body.theme-light .slide-preview-label { font-size: 10.5px; color: var(--faire-sub); letter-spacing: 0.7px; text-transform: uppercase; }

/* Notes */
body.theme-light .speaker-notes-container {
  flex: 1; min-height: 0;
  margin: 0 18px;
  background: var(--faire-surface);
  border: 1px solid var(--faire-border);
  border-radius: var(--faire-radius);
  display: flex; flex-direction: column; overflow: hidden;
}
body.theme-light .notes-zoom-controls {
  position: static; display: flex; visibility: visible;
  padding: 6px 8px 6px 14px; background: var(--faire-warm);
  border-bottom: 1px solid var(--faire-border); gap: 6px; align-items: center;
  margin-top: 0;
}
body.theme-light .notes-zoom-controls::before {
  content: 'Speaker notes'; font-size: 10.5px; letter-spacing: 0.7px;
  text-transform: uppercase; color: var(--faire-sub); flex: 1;
}
body.theme-light .notes-zoom-btn {
  background: var(--faire-surface); border: 1px solid var(--faire-border);
  border-radius: var(--faire-radius); padding: 0; width: 32px; height: 30px;
  display: inline-flex; align-items: center; justify-content: center; color: var(--faire-text);
  font-size: 13px; font-weight: 400;
}
body.theme-light .notes-zoom-btn:hover { background: var(--faire-warm); color: var(--faire-text); border-color: var(--faire-border); }
body.theme-light #notes-zoom-readout { font-family: var(--faire-font-mono); font-size: 11px; color: var(--faire-sub); padding: 0 8px; }

body.theme-light .speaker-notes-content-wrapper { padding: 16px 18px; flex: 1; overflow-y: auto; border-radius: 0; border: none; background: transparent; }
body.theme-light .speaker-notes-content { font-family: var(--faire-font-serif); font-size: 19px; line-height: 30px; color: var(--faire-text); }

/* Controls */
body.theme-light .remote-controls { padding: 14px 18px; display: flex; gap: 10px; }
body.theme-light .remote-btn {
  flex: 1; height: 72px; border-radius: var(--faire-radius);
  display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 2px;
  font-family: var(--faire-font-sans); font-size: 11px; letter-spacing: 0.5px; text-transform: uppercase;
}
body.theme-light .remote-btn-prev { background: var(--faire-surface); border: 1px solid var(--faire-border); color: var(--faire-text); }
body.theme-light .remote-btn-next { background: var(--faire-text); border: 1px solid var(--faire-text); color: var(--faire-surface); }

/* Bottom nav */
body.theme-light .bottom-tabs { border-top: 1px solid var(--faire-border); background: var(--faire-surface); display: flex; padding: 6px 0 12px; }
body.theme-light .bottom-tabs .tab-btn {
  flex: 1; background: transparent; border: none; border-radius: 0; padding: 6px 0;
  display: flex; flex-direction: column; align-items: center; gap: 2px;
  color: var(--faire-sub); font-size: 10px; letter-spacing: 0.2px; font-weight: 400;
}
body.theme-light .bottom-tabs .tab-btn.active { color: var(--faire-text); font-weight: 500; background: transparent; }
```

## 8. Bottom-tab nav markup

Because existing JS tab-switching reads `.tab-btn[data-tab]`, the simplest approach is:

- Leave the original top `.tabs` block in the DOM (required by other themes) but hide it in `theme-light` via `display: none`.
- Append a second `<nav class="bottom-tabs">` with three **mirror** `<button>`s, each with `class="tab-btn"`, `data-tab="remote|controls|settings"`, and an icon SVG.
- If the existing handler scopes its selectors (e.g. `.tabs .tab-btn`), broaden to `.tab-btn` so both sets sync active state:

```js
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.classList.toggle('active', btn.dataset.tab === tab);
});
```

## 9. Fonts

Inject at the top of the remote `<head>` template:

```html
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Lora:ital,wght@0,400;0,500;1,400&display=swap" rel="stylesheet">
```

For offline / restricted-tunnel scenarios, self-host WOFF2 files under `/fonts/*` with `@font-face` fallback. System fonts are acceptable degradation.

## 10. Manual test checklist (smoke)

1. Set `web-ui-theme=light`, open the remote on an iPhone-sized viewport. Expect V2-C layout.
2. Disable Stagetimer — container hidden.
3. Enable Stagetimer with no API key — container stays hidden.
4. Configure a valid room — container appears. Cycle states:
   - idle → neutral warm
   - running > 60s → green
   - 60s ≥ remaining > 15s → amber
   - ≤ 15s → red
   - remainingMs < 0 → **overtime** (dark red inverse)
   - error → red
5. Push a message from Stagetimer.io — shows tone-matched below the clock row. Card grows; notes card below does not jump.
6. Tap `Next slide` — verify `POST /api/slide-next` fires.
7. Tap `Previous` — verify `POST /api/slide-prev` fires.
8. Tap zoom +/− — `#speaker-notes-content` font-size updates, `#notes-zoom-readout` reflects new value, persists across refresh.
9. Tap scroll ↑/↓ — presenter-machine notes display scrolls (verify `/api/notes-scroll-up` / `/api/notes-scroll-down` fires).
10. Bottom tabs switch to Controls and Settings tabs (unmodified content).
11. In restricted-tunnel mode (`webUiRestrictedTunnelClient=true`), Settings tab button hidden in both top and bottom nav.
12. Other themes (`dark`, `max`, `touch`, `thumb`) unchanged.

Special cases to verify overtime:
- Let a timer run past zero with both the REST fallback path and the socket.io live path active; both paths must enter `.overtime` because both updater functions were patched.

## 11. Out of scope

- Electron preferences window (`index.html`, `renderer.js`).
- Speaker Notes window.
- Other themes (`dark`, `max`, `touch`, `thumb`).
- Companion module, Chrome extension.
- Preferences schema or IPC.
- Restructuring `main.js` beyond the template-literal edits described.

## 12. Suggested commit breakdown

1. `feat(web-ui): introduce Faire design tokens for theme-light`
2. `feat(web-ui): restructure Remote tab DOM for V2-C layout (preserves ids)`
3. `feat(web-ui): rewrite theme-light CSS (stagetimer, notes, controls, bottom tabs)`
4. `feat(stagetimer): add overtime state for remainingMs < 0 in both updaters`
5. `feat(web-ui): add notes zoom px readout`
6. `chore(web-ui): load Inter + Lora via Google Fonts link`
