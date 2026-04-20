# Web Remote V2-C — Implementation Plan

Target: refresh the web UI served by the Electron main process (the remote page at `GET /`) to the **V2-C "Stage-ready, notes-expanded"** design. Visual reference: `docs/desktop-ui-mock/Web Remote UI.html` (V2-C artboard in the canvas).

## 0. Guardrails

- **Do not break the REST API.** Endpoints (`/api/slide-prev`, `/api/slide-next`, `/api/notes-scroll-up`, `/api/notes-scroll-down`, `/api/get-stagetimer-status`, socket.io events, etc.) are consumed by the Companion module and Chrome extension. Do not rename them.
- **Do not change preferences schema.** `web-ui-theme`, `web-ui-logo-*`, `web-ui-custom-css-*`, `stagetimer-*`, `default-notes-zoom-steps` all stay. This is a CSS/HTML refactor inside the `theme-light` branch (or a new `theme-faire` option — see §1).
- **Do not change Settings / Controls tabs content.** Only re-style them to match. Only the **Remote** tab is restructured.
- **Preserve every DOM id.** The existing event wiring in `main.js` (~lines 7000–8100) binds by id:
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

**Option B:** add a new theme option `faire` alongside the existing five. Adds a row to the `web-ui-theme` `<select>` in `index.html:385` and in the main-process HTML (`main.js` around line 388 in that second `<option>` list). Default stays `light`; opt-in migration.

Plan below assumes **Option A**.

## 2. Files to edit

All changes live in three blocks inside `main.js` (the remote page is a giant template literal in the `buildWebRemoteHTML` function, ~line 6150 onward):

1. **CSS block 1** — base component styles, ~lines 6180–6340 (`.notes-zoom-controls`, `.stagetimer-container`, etc.)
2. **CSS block 2** — theme overrides, ~lines 6325–6410 (the `body.theme-light` block specifically)
3. **HTML block** — Remote tab markup, ~lines 6440–6520

No changes needed in: `index.html`, `renderer.js`, `preload.js`, preferences, IPC.

## 3. Design tokens

Add these CSS custom properties at the top of the `<style>` block so all new rules key off them (and a later dark theme can override by flipping tokens):

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

Load Google Fonts `Inter` (400/500/600) and `Lora` (400/500 + 400italic) via `<link>` at the top of the remote HTML. Keep a system-font fallback so offline / locked-down tunnels still render.

## 4. Remote-tab DOM restructure (V2-C)

Rewrite the `<div id="tab-remote" class="tab-content active">` block so the order top→bottom is:

1. **Header row** (new wrapper `<div class="remote-header-compact">`)
   - Status dot (reuse `.system-icon` position) + `machineName` in `--faire-text` 500
   - Slide counter "3 / 24" right-aligned, `--faire-sub`
   - `#notes-toggle-btn` and `#previews-toggle-btn` move here as icon-only 32×32 buttons (keep ids and click handlers).
2. **`#stagetimer-container`** (V2-C expects this always-mounted so no layout shift; CSS handles hidden state — see §5).
3. **Slide strip** (replaces `.slide-previews-grid`)
   - Same grid ids/structure but each card becomes `[thumbnail 96×54] [Label / slide title]` horizontally.
   - The thumb image keeps `#slide-preview-current-img` / `#slide-preview-next-img`.
   - New title element per card — for now reuse the existing `#slide-preview-current-label` / `#slide-preview-next-label` text nodes; if server data includes slide title later, populate there.
3. **Notes card** (`#speaker-notes-container`, made flex:1)
   - Toolbar row: "Speaker notes · slide N" on the left; scroll cluster (↑/↓) and zoom cluster (−/px-readout/+) on the right.
   - Keep all four buttons with their existing ids; replace text labels ("Scroll Up") with SVG icons + `aria-label`.
   - Add `<span id="notes-zoom-readout">18px</span>` between `#notes-zoom-out` and `#notes-zoom-in`; update it in the existing zoom handlers (~lines 7191, 7200 of main.js) by writing the computed `font-size` back.
   - Wrapper body stays `#speaker-notes-content-wrapper > #speaker-notes-content` — only the typography changes (Lora 19/30).
4. **`#remote-controls`** — two big buttons, `Previous` secondary and `Next slide` primary, 72px tall.
5. **Bottom tab bar** (replaces the top `.tabs` row for `theme-light` only). Three `.tab-btn` buttons stay with their `data-tab` attributes, but re-styled as icon + label in a `<nav class="bottom-tabs">` at the bottom of the `.container`. Active state uses `--faire-text`, inactive `--faire-sub`.

## 5. Stagetimer — state-driven colors

The existing code in `updateStagetimerDisplay` (main.js ~L7686 and ~L8068) already sets one of:
- `.stagetimer-container` (idle/unknown)
- `.stagetimer-container running`
- `.stagetimer-container warning` (remainingSeconds ≤ 60)
- `.stagetimer-container critical` (remainingSeconds ≤ 15)
- `.stagetimer-container error`
- `.stagetimer-container disabled`

**Minimal JS change:** add one more state — overtime — when `remainingMs < 0`. In both update functions:

```js
if (data.remainingMs !== undefined) {
  const remainingSeconds = Math.floor(data.remainingMs / 1000);
  if (remainingSeconds < 0)       container.className = 'stagetimer-container overtime';
  else if (remainingSeconds <= 15) container.className = 'stagetimer-container critical';
  else if (remainingSeconds <= 60) container.className = 'stagetimer-container warning';
}
```

Adjust thresholds to taste — 15s/60s are today's; plan advises 30s/20% but keep current values unless product confirms.

**CSS rewrite** for the `theme-light` branch (drop the gradient):

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

The existing layout has the clock centered and the messages absolutely-positioned (explains the fixed 160px height). In the new layout, header row + messages flow vertically inside the card — no fixed height, no layout shift because the card simply grows when a message is pushed.

**Rework the markup** inside `#stagetimer-container` so it's:

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

The existing JS writes `textContent` into each of those three ids — no JS refactor needed beyond adding the `overtime` branch. `stagetimer-dot` is purely decorative (CSS).

## 6. Notes zoom readout

Existing handlers (main.js ~L7191/7200) change `speakerNotesContent.style.fontSize`. After each change, also:

```js
const readout = document.getElementById('notes-zoom-readout');
if (readout) {
  const px = parseInt(getComputedStyle(speakerNotesContent).fontSize, 10);
  readout.textContent = px + 'px';
}
```

Call it once on init (after restoring persisted zoom level from localStorage).

## 7. CSS for the Remote tab (theme-light branch)

New block (replaces all current `body.theme-light .*` rules for Remote-tab elements):

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
- Append a second `<nav class="bottom-tabs">` containing three **mirror** `<button>`s, each with `class="tab-btn"`, `data-tab="remote|controls|settings"`, and an icon SVG.
- The existing event handler (`.tab-btn` click → switches `.active` on tab-buttons + tab-content) already handles `document.querySelectorAll('.tab-btn')` — both sets will sync.

If the existing handler uses `document.querySelector('.tab-btn[data-tab=...]')` (singular), broaden it:

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

For offline/restricted-tunnel scenarios, also ship self-hosted WOFF2 files under `/fonts/*` and add `@font-face` declarations as fallback (or accept system fonts — not a blocker).

## 10. Manual test checklist (smoke)

After implementing:

1. Set `web-ui-theme=light`, open the remote on an iPhone-sized viewport. Expect the V2-C layout.
2. Disable Stagetimer — container hidden (`display: none`).
3. Enable Stagetimer with no API key — container stays hidden.
4. Configure a valid room — container appears. Watch it cycle:
   - idle / no running timer → neutral warm
   - running with >60s left → green
   - 60s ≥ remaining > 15s → amber
   - ≤ 15s → red
   - remainingMs < 0 → overtime (dark red inverse)
   - error → red (uses critical tone)
5. Push a message from Stagetimer.io — shows as tone-matched block below the clock row. Card grows, no layout shift in surrounding elements (verify notes card doesn't jump).
6. Tap `Next slide` — advances (verify `POST /api/slide-next` fires).
7. Tap `Previous` — retreats.
8. Tap zoom +/− — `#speaker-notes-content` font-size updates, `#notes-zoom-readout` shows new value, persists in localStorage across refresh.
9. Tap scroll ↑/↓ — presenter-machine notes display scrolls (verify `/api/notes-scroll-up` / `/api/notes-scroll-down` fires).
10. Tap bottom-tab `Controls` — switches to Controls tab (unmodified content).
11. Tap bottom-tab `Settings` — switches to Settings tab.
12. In restricted-tunnel mode (`webUiRestrictedTunnelClient=true`), verify Settings tab button is hidden (same condition as existing code).
13. Other themes (`dark`, `max`, `touch`, `thumb`) — verify unchanged.

## 11. Out of scope

- Any changes to the desktop Electron preferences window (`index.html` / `renderer.js`).
- Any changes to the Speaker Notes window (`speaker-notes-window.html` or equivalent).
- Any changes to other themes (`dark`, `max`, `touch`, `thumb`) — they should continue to work exactly as before because the new tokens and rules are scoped to `body.theme-light`.
- Companion module and Chrome extension.
- Adding new preferences.
- Restructuring `main.js` beyond the template-literal edits described.

## 12. Suggested commit breakdown

1. `feat(web-ui): introduce Faire design tokens for theme-light`
2. `feat(web-ui): restructure Remote tab DOM for V2-C layout (preserves ids)`
3. `feat(web-ui): rewrite theme-light CSS (stagetimer, notes, controls, bottom tabs)`
4. `feat(stagetimer): add overtime state for remainingMs < 0`
5. `feat(web-ui): add notes zoom px readout`
6. `chore(web-ui): load Inter + Lora via Google Fonts link`
