# Desktop UI Redesign — Implementation Plan

> **Scope:** Redesign of the **main Electron settings/dashboard window only** (`index.html` + `renderer.js` + `styles.css`). Do **not** touch the presentation window, the speaker-notes/presenter window, or any web-remote UI. No changes to `main.js`, IPC, `preload.js`, API endpoints, preferences schema, or Google Slides rendering behavior.
>
> **Reference mock:** `docs/desktop-ui-mock/GSlide Controller Desktop.html` (self-contained React/Babel prototype showing the intended structure, spacing, and copy). The mock is static — it does not call any IPC. Your job is to port its **IA, visual language, and interaction model** onto the existing renderer, keeping all real backend wiring intact.

---

## 1. What changes, in one paragraph

Today the desktop window is one long vertical scroll of `<section class="card">` blocks — Google Account, Primary/Backup, Speaker notes, Ports, URLs, WAN/Tunnel, Displays, Presets, Stagetimer, Appearance, Debug, Logs, Crash reports — every setting visible at all times. The redesign introduces a **left sidebar** with grouped navigation and promotes a new **Dashboard** view that leads with the share URLs and a launch checklist. Each existing `<section>` is reassigned to a tab; only one tab renders at a time. A persistent **top status bar** surfaces sign-in state, API port, tunnel status, and primary/backup mode so the window can be minimized to a background role during shows. Faire-inspired visual treatment: warm-gray palette, Inter body, Lora serif for page titles, 1px borders, no drop shadows.

---

## 2. Information architecture (final)

Sidebar, in order:

**Overview**
- **Dashboard** — share URLs, WAN/tunnel toggle, launch checklist, "Open test presentation" / "Open a preset" CTAs, inline machine-name rename.
- **Speaker notes** — current-slide notes with the live-slide indicator, Refresh, Copy to clipboard.

**Setup**
- **Account** — Google sign-in status, sign out.
- **Network** — Share URLs (API + web), **Ports** (API, Web UI, **HTTPS toggle is here, adjacent to the ports** — cert/key pickers reveal when HTTPS is on), WAN/Cloudflare tunnel, PIN.
- **Monitors** — presentation display, notes display, notes layout, default notes zoom, "Relaunch notes" button.
- **Presets** — preset list (label + URL rows), save / reload.
- **Web remote** — theme, brand logo, custom CSS, Stagetimer.io.

**Other**
- **Advanced** — Primary/Backup mode + backup machines, controller IP allowlist, GPU mode, native fullscreen. **(HTTPS is no longer here — it moved to Network.)**
- **Logs** — live log tail, verbosity toggles, crash-reports block, Save `.txt`.

Top status bar (persistent, above the content area, below traffic lights):
- Machine name (left)
- Sign-in dot + email
- API port (`:9595`)
- Tunnel state (`LAN only` / `Tunnel active`)
- Mode (`Standalone` / `Primary` / `Backup`)

---

## 3. Field-by-field mapping from old → new

Every element ID and DOM node in `index.html` maps to exactly one new location. **Do not delete any `id`** — `renderer.js` and `preload.js` query by ID and removing one silently breaks settings persistence. Move the node, don't re-create it.

| Existing `<section class="card">` h2 | Contains IDs | New tab | Notes |
|---|---|---|---|
| Google Account | `signin-btn` | **Account** | Unchanged content; new shell |
| Primary/Backup Configuration | `mode-primary/backup/standalone`, `backup-config`, `backup-port`, `backup-ip-list`, `add-backup-ip` | **Advanced** → "Primary / Backup" panel | |
| Speaker notes (current slide) | `speaker-notes-capture`, `notes-encoding-warning-desktop`, `speaker-notes-refresh`, `speaker-notes-copy` | **Speaker notes** (top-level tab) | Keep the encoding-warning element verbatim |
| Network Ports | `api-port`, `web-ui-port`, `web-ui-use-https`, `web-ui-https-cert-group`, `web-ui-https-key-group`, `web-ui-cert-*`, `web-ui-key-*` | **Network** → "Ports" panel | HTTPS toggle lives directly under Web UI port; cert/key rows still toggle based on `web-ui-use-https` |
| (URLs block) | `api-urls`, `web-ui-urls` | **Dashboard** ("Share URLs" panel) **and** **Network** ("Share URLs" panel) | Render same lists in both — they are cheap, and operators look in both places. See §5 for how. |
| WAN / Tunnel | `wan-enabled`, `wan-status-row`, `wan-url-display`, `wan-copy-btn`, `wan-tunnel-pin-*`, `wan-pin-scope` | **Network** → "Share over internet" panel | Also surface the `wan-enabled` toggle + `wan-url-display` (read-only) on Dashboard |
| Monitors / Displays | `presentation-display`, `notes-display`, `notes-layout`, `btn-relaunch-speaker-notes`, `default-notes-zoom-steps` | **Monitors** | |
| Presentation rendering | `presentation-gpu-mode`, `presentation-native-fullscreen`, `save-presentation-rendering-btn` | **Advanced** → "Presentation rendering" panel | |
| Machine name | `machine-name` | **Dashboard** (primary) **and** **Account** (secondary) | See §4 — Dashboard shows an inline-editable heading bound to the same input. |
| Controller IP allowlist | `controller-ip-list`, `add-controller-ip` | **Advanced** → "Controller allowlist" panel | |
| Presets | `preset-list`, `add-preset`, `save-presets-btn`, `load-presets-btn` | **Presets** | |
| Stagetimer | `stagetimer-room-id`, `stagetimer-api-key`, `stagetimer-enabled`, `stagetimer-visible`, `save-stagetimer-btn`, `load-stagetimer-btn` | **Web remote** → "Stagetimer.io" panel | |
| Web UI Appearance | `web-ui-theme`, `web-ui-logo-*`, `web-ui-custom-css-*`, `save-web-ui-appearance-btn`, `web-ui-download-css-template` | **Web remote** → "Appearance" panel | |
| Test button | `test-btn` | **Dashboard** CTA row | |
| Debug / Crash info | `crash-info-block`, `crash-reports-dir`, `crash-dumps-dir`, `last-crash-time`, `open-crash-reports-folder-btn`, `debug-logs-console`, `debug-logs-clear`, `debug-logs-save`, `verbose-logging`, `web-ui-debug-console-enabled` | **Logs** | |
| Status message | `status-message` | Stays global — render it as a toast anchored to the bottom-right of the window frame, outside any tab |
| Build number | `build-number` | Stays global — render in the sidebar footer |

---

## 4. Two small UX features that are new (not in v1.9.2)

### 4a. Inline machine-name rename on Dashboard

The web UI lets operators click the system name to rename in place. Port the same pattern to Dashboard:

- Render `s.machineName` as a large (32px Lora) button/heading. On hover, a pencil glyph fades in.
- Click → swap to an `<input>` that inherits the heading's typography.
- **Enter** or **blur** commits; **Escape** reverts.
- On commit: write through to the existing `#machine-name` text input (the one in the old code), then dispatch a `change` event (that is the only event `renderer.js` listens for on this input — no `input` event listener exists). **Do not add a second preferences round-trip** — we're piggybacking on the existing bind so preferences persistence is untouched.
- Keep the canonical `#machine-name` input mounted (hidden or on the Account tab). The Dashboard heading is a view over it.

### 4b. HTTPS moved next to Ports

Previously HTTPS sat in its own section far from the port configuration. New location: **Network → Ports panel**, directly under "Web UI port" as a toggle row labeled "Serve over HTTPS". The cert/key file-picker rows (`web-ui-https-cert-group`, `web-ui-https-key-group`) remain, and continue to show/hide based on `web-ui-use-https`. No behavior change, just proximity.

---

## 5. Implementation approach

Two viable routes. **Pick Route A unless the team explicitly wants a framework migration.**

### Route A — Stay vanilla, restructure DOM and CSS (recommended)

1. **Rewrite `index.html` shell** to this outline:

   ```html
   <div class="app">
     <aside class="sidebar">…nav items…<footer>build #…</footer></aside>
     <div class="main">
       <header class="status-bar">…machine name, sign-in dot, port, tunnel, mode…</header>
       <div class="content">
         <section data-tab="dashboard">…</section>
         <section data-tab="notes" hidden>…</section>
         <section data-tab="account" hidden>…</section>
         <section data-tab="network" hidden>…</section>
         <section data-tab="monitors" hidden>…</section>
         <section data-tab="presets" hidden>…</section>
         <section data-tab="remote" hidden>…</section>
         <section data-tab="advanced" hidden>…</section>
         <section data-tab="logs" hidden>…</section>
       </div>
     </div>
   </div>
   ```

2. **Move every existing `<section class="card">` into the appropriate `<section data-tab>`**, preserving inner markup and `id` attributes exactly. You are physically cutting and pasting, not re-authoring. The one rewrite is the Dashboard, which is new composition (share URLs panel + checklist + CTA row + the new inline-rename heading).

3. **Nav switching** is a 30-line vanilla controller in `renderer.js`:
   - Click handler on `.sidebar [data-target]` toggles `hidden` on the matching `section[data-tab]`, adds an `.active` class to the clicked link, and persists the last-selected tab to `localStorage` (`desktop-ui:activeTab`).
   - Read it back on load so refresh returns you to the same view.

4. **Rewrite `styles.css`** to the Faire-style tokens from the mock. Copy these values verbatim — they are already tuned:

   ```css
   :root {
     --bg: #ffffff;
     --surface-2: #fbf8f6;   /* warm neutral, used for speaker-notes surface and sidebar hover */
     --surface-3: #f3f0ed;
     --border:   #dfe0e1;
     --text:     #333333;
     --text-subdued: #757575;
     --text-disabled: #b5a998;
     --ok:   #49694c;
     --warn: #907c3a;
     --bad:  #921100;
     --font-sans: 'Inter', -apple-system, BlinkMacSystemFont, 'SF Pro Text', sans-serif;
     --font-serif: 'Lora', Georgia, serif;
     --font-mono: ui-monospace, SFMono-Regular, Menlo, monospace;
     --radius: 4px;
   }
   ```

   Replace the existing `.card` / `.btn` / `.input-field` / `.select-input` / `.form-group` styles with the versions demonstrated in the mock. Panels: 1px solid `var(--border)`, 4px radius, no box-shadow. Buttons: near-black primary (`#333` bg, `#fff` text), secondary is transparent with a 1px border. Page titles: Lora 28–32px. Body: Inter 13px.

5. **Dashboard share-URL duplication.** Rather than cloning DOM (fragile — two `#api-urls`), extract the URL-list renderer in `renderer.js` into a helper:

   ```js
   function renderUrlLists() {
     renderUrlListInto('#dashboard-api-urls', apiUrls);
     renderUrlListInto('#dashboard-web-ui-urls', webUiUrls);
     renderUrlListInto('#api-urls',              apiUrls);
     renderUrlListInto('#web-ui-urls',           webUiUrls);
   }
   ```

   Then call `renderUrlLists()` everywhere the code currently writes to `#api-urls` / `#web-ui-urls`. IDs on the Dashboard side are prefixed to avoid collision.

6. **Top status bar** is read-only — it reflects state the renderer already tracks. Add four small DOM nodes and update them inside the existing state-change paths: `isSignedIn`, port inputs' change listeners, `wan-enabled` toggle, the primary/backup radio group. No new IPC.

7. **Inline rename** (see §4a). The `#machine-name` `<input>` stays mounted on the Account tab. On Dashboard render, read its value into the heading and wire the edit-in-place widget to write back to the input and dispatch a `change` event — that is the only event `renderer.js` listens for on this input.

8. **Nothing else in `renderer.js` should change.** Keep all IPC calls, event listeners, and state in place.

### Route B — React rewrite

Only take this on if there's appetite for a real framework migration. The mock is React/Babel-in-browser for prototyping speed, not a production pattern. If you go this way, use Vite + React 18, port each tab as a component, and wrap the existing `window.electronAPI` in a thin hooks layer. Scope creep risk is high — do not combine with this plan.

---

## 6. Files touched

- `index.html` — rewritten shell + section relocations. All `id`s preserved.
- `styles.css` — new tokens and component styles (full rewrite is fine; the mock is the reference).
- `renderer.js` — add: tab switcher, URL-list helper, top-status-bar updaters, inline-rename widget. **Do not** edit any code that talks to `window.electronAPI`, Google auth, preferences, WAN, stagetimer, backup failover, or the speaker-notes pipeline.
- No changes to: `main.js`, `preload.js`, `package.json`, build scripts, any file under `companion-module-gslide-opener/`, `gslide-chrome-extension/`, `gslide-websocket-addon/`, `sidecar/`, or `resources/`.

---

## 7. Acceptance checklist

Smoke test after the port — every one of these must still work:

- [ ] Sign in with Google, sign out, session persists across app restart.
- [ ] Change API port, restart app, new port is honored.
- [ ] Toggle HTTPS (now on Network tab), pick cert + key files, web UI is reachable over `https://`.
- [ ] Toggle WAN tunnel, URL appears on Dashboard AND Network; copy button works.
- [ ] Set, save, and reload presets. Companion preset actions 1/2/3 still fire.
- [ ] Switch primary/backup modes; backup IP list reveals/hides; backup status pills on Primary reflect real reachability.
- [ ] Speaker notes tab shows current slide's notes; encoding warning appears when triggered; Copy works.
- [ ] Monitors tab changes presentation display; "Relaunch notes" works.
- [ ] Logs tab streams live entries; Save `.txt` writes a file; Clear empties.
- [ ] Crash-reports folder opens.
- [ ] Stagetimer room/API-key save, enable/disable works, visible toggle works.
- [ ] Web-UI theme, logo, and custom CSS save; Download CSS template works.
- [ ] Controller allowlist save works; bad IP still rejects requests.
- [ ] Inline rename on Dashboard updates `#machine-name`, persists, and the old Account-tab input shows the same value.
- [ ] Refresh / restart returns to the last-selected tab.
- [ ] No console errors in the renderer on cold start or when switching tabs.

---

## 8. What is **explicitly out of scope**

- Presentation window styling, chrome, or behavior.
- Speaker-notes/presenter window styling or behavior. (The **Speaker notes tab** in the desktop window is in scope; the **presenter window** is not.)
- Web-remote (browser) UI — any files served from the local HTTP server to remote clients.
- Companion module (`companion-module-gslide-opener/`).
- Chrome extension, websocket add-on, sidecar, cloudflared resources.
- Any change to preferences JSON shape, IPC channels, or API endpoints.
- Dark mode. Ship light-only for this pass.
- i18n. English only.

---

## 9. Reference

The visual target is a self-contained React/Babel prototype at:

```
docs/desktop-ui-mock/GSlide Controller Desktop.html
```

Open it in a browser to see the exact spacing, type scale, hover states, and copy. The mock uses mock state and has no IPC — treat it as a visual contract, not as code to port.

