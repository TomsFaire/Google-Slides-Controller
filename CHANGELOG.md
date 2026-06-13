# Changelog

All notable changes to Google Slides Opener are documented here.

## [2.3.5] - 2026-06-12

### Added
- **Stage timer overlay on notes monitor** — Frameless, always-on-top countdown clock on the notes display, connected to StageTimer.io via Socket.io (reuses existing room ID and API key). Color-codes through idle → running → warning → critical → overtime. Controls in desktop settings (enable, position, size %) and web remote Settings (show/hide, position, size). HTTP API: `POST /api/show-stage-timer-overlay`, `POST /api/hide-stage-timer-overlay`, `POST /api/update-stage-timer-overlay-settings`; `GET /api/status` adds `stageTimerOverlayEnabled`, `stageTimerOverlayPosition`, `stageTimerOverlaySize`. IPC: `window.electronAPI.stageTimerOverlay.{show,hide,getStatus,updateSettings}`.

### Fixed
- **Display labels** — Primary display detection uses `screen.getPrimaryDisplay().id` in IPC and HTTP handlers; labels normalized to `Display N (W×H)`. Renderer no longer stuck on "Loading…" when display enumeration fails.
- **StageTimer visibility** — Unchecking "Show timer on Remote tab" and saving now hides the embedded StageTimer widget (respects `stagetimerVisible`, not just API key presence).
- **Keyboard shortcut hint** — Tooltip scoped to the Remote tab only; hidden when switching to Controls or Settings.

### Companion module
- Ship **companion-module-gslide-opener v1.7.0** — Stage timer overlay: 4 actions (`show_stage_timer_overlay`, `hide_stage_timer_overlay`, `set_stage_timer_overlay_position`, `set_stage_timer_overlay_size`), 3 variables, 1 feedback (`stage_timer_overlay_active`).

### Build
- **Version 2.3.5**, **build 82**.

---

## [2.3.4] - 2026-05-15

### Added
- **WaveShare RS232/485/422 TO POE ETH (B) support for PerfectCue** — Per-port adapter selection (DSAN vs WaveShare), WaveShare byte mapping and RS485 debounce, keep-alive presets, and setup/testing docs (`docs/waveshare-perfectcue-setup.md`, `docs/waveshare-perfectcue-testing.md`).

### Build
- **Version 2.3.4**, **build 81**.

---

## [2.3.3] - 2026-05-14

### Added
- **Configurable keyboard shortcut preset** — Operators can now choose between three presets from the Electron desktop app's Web Remote tab or the web remote's Settings tab:
  - `Cmd/Ctrl + Arrow` (original, default)
  - `Alt/Option + Arrow`
  - `Cmd/Ctrl + Shift + Arrow` (safest — avoids browser back/forward conflict)
- **Default-enabled toggle** — Admins can pre-enable keyboard shortcuts for all new connections without requiring users to tap the toggle.
- Both settings persist in `preferences.json` and are applied on web remote load via injected template globals.
- Preset changes in the web Settings tab take effect immediately without a page reload.

## [2.3.2] - 2026-05-14

### Added
- **Keyboard shortcuts for web remote** – tap the keyboard icon in the Remote tab header to enable `Cmd+←/→` (previous/next slide) and `Cmd+↑/↓` (scroll speaker notes). On Windows/Linux use `Ctrl`. The toggle persists across page reloads via `localStorage`. Shortcuts are suppressed when an input, textarea, or select element has focus.

### Build
- **Version 2.3.2**, **build 78**.

---

## [2.3.0] - 2026-05-06

### Added
- **Open arbitrary URLs on the presentation display** – `POST /api/open-url` opens any `http`/`https` URL in a dedicated Electron session partition (`persist:generic`) so cookies stay isolated from Google Slides and Slido. Requires `allowArbitraryUrl` preference to be enabled.
- **LAN-only enforcement** – The "Open URL" section in the web remote is server-side rendered only for LAN clients (hidden when connecting over the WAN/Cloudflare tunnel). The web UI proxy blocks tunnel clients from calling `/api/open-url` before the request reaches the API server.
- **Configurable background color** – `genericUrlBackgroundColor` preference (hex, default `#000000`) sets the BrowserWindow background shown before a page loads or behind transparent content.
- **Desktop app toggle** – Advanced → Web Remote Features panel: "Allow opening arbitrary URLs (LAN only)" checkbox and background color picker (with live hex input sync).
- **`contentKind` value `generic-url`** – `GET /api/status` returns `generic-url` when an arbitrary URL is showing; crash recovery and `/api/reload-presentation` handle the new kind.
- **Web remote "Open URL" section** – URL input + button with Enter-key support; only visible when the feature is enabled and the client is on LAN.

### Companion module
- Ship **companion-module-gslide-opener v1.6.0** – **Open URL** action (`open_url`) with variable support. Companion connects directly to the API on port 9595 (LAN-only by network topology) so the proxy-layer tunnel gate does not apply.

### Build
- **Version 2.3.0**, **build 76**.

## [2.2.0] - 2026-04-28

### Added
- **Slido on the presentation display** – `POST /api/open-slido` opens an https `*.sli.do` URL (e.g. wall) in a dedicated Electron session partition (`persist:slido`) so Okta/SSO cookies stay separate from Google Slides. Popups use default window behavior (no speaker-notes overrides) for MFA. Controller allowlist and backup `sendToBackups('/api/open-slido')` match other mutating APIs.
- **`contentKind` in `GET /api/status`** – `slides` or `slido` so automation knows what is showing.
- **Web UI (Controls tab)** – Slido URL field and **Open Slido** button; Enter submits.

### Changed
- **`POST /api/reload-presentation`** – Reloads Slido by reopening the same URL when `contentKind` is Slido (Slides behavior unchanged).

### Companion module
- Ship **companion-module-gslide-opener v1.5.0** – **Open Slido** action and **`content_kind`** variable.

### Build
- **Version 2.2.0**, **build 71**.

## [2.1.1] - 2026-04-28

### Security / networking
- **PerfectCue TCP** – Connections are allowed only if the client address passes the same controller IP allowlist used for the HTTP API (`src/perfectcue-server.js`).

### Build / CI
- **Companion module artifact** – Workflow packages **`companion-module-gslide-opener`** from this repository (not the upstream Bitfocus fork).

### Tests
- **PerfectCue** – Additional tests for port normalization and TCP dispatch gating.

## [2.0.1] - 2026-04-23

### Fixed
- **Web remote light theme – speaker notes stuck at "Loading notes..."** – In the light (minimalist) theme the speaker notes panel is always visible via CSS, so users never needed to click the toggle button. But notes polling only started on toggle click, leaving the panel permanently stuck at the placeholder text. Notes now auto-start on page load when the light theme is active.

## [2.0.0] - 2026-04-20

### Desktop app (Electron settings)
- **New shell** – Sidebar navigation plus tabbed content for a clearer settings workflow.
- **Visual system** – Faire-inspired design tokens, typography, cards, and button hierarchy for a more polished, consistent window.
- **Quality-of-life** – Tab switching and URL handling improvements, status bar refinements, and inline rename flows where presets and lists are edited.
- **Accessibility** – Stronger focus states, nav semantics, and small interaction fixes from review.

### Web remote (browser UI served by the app)
- **Light theme (V2-C)** – Full-height remote layout, warm Faire-style surfaces, bottom tab bar for Remote / Controls / Settings, wider layout on tablet and desktop, and safer scrolling (bottom nav moved outside `overflow: hidden` so it is not clipped).
- **All themes** – Shared structure for Controls and Settings (section cards, primary/secondary buttons, inputs) so every theme feels like the same product; per-theme color tokens for sections, accents, and StageTimer states.
- **StageTimer** – Non-light themes use a compact flat layout with clear idle / running / warning / critical / overtime / disabled palettes and tabular clock digits.
- **Presets in the browser** – Clearer empty state (link into Settings), launch row layout, and warning callouts where appropriate.
- **Reliability** – Root URL works with query strings; HTML responses use `no-store` where appropriate; inactive tab panels stay hidden so Remote content cannot leak onto Controls or Settings on iPad-sized layouts.
- **Corner radius** – Buttons, tabs, panels, and preview chrome use the same radius language as the light theme (`4px` controls, `10px` settings cards, `2px` slide preview thumbnails) across original, dark, max, touch, and thumb themes.

### Documentation
- **README** — New `docs/images/` gallery for v2 desktop settings and light-theme web remote; maintainer script `yarn capture:readme-screenshots` (temporarily uses Web UI port **8765**, then restores preferences).

### Build
- **Version 2.0.0**, **build 70**.

## [1.9.12] - 2026-04-15

### Fixed / rendering
- **Google Slides transparent PNG / layer compositing** – Upgraded **Electron to 33.x** (newer Chromium, closer to Chrome). Added **Settings → Monitor Setup → Presentation GPU mode** (`default`, ANGLE **Metal** / **OpenGL**, **SwiftShader**, or **disable GPU**) applied at startup via `app.commandLine` / `disableHardwareAcceleration()`; **restart required** after changing GPU mode. Optional **macOS native fullscreen** for the slide window (diagnostic compositor path vs simple fullscreen).

### Added
- **`GET /api/status`** – `runtime` (`chrome`, `electron`, `node`), `presentationGpuMode`, `presentationNativeFullscreen`. **`get-build-info`** includes Chromium/Electron/Node versions for the footer.
- **`yarn smoke:slides-gpu`** – Minimal Electron window to reproduce presenter rendering; optional `GSLIDE_GPU_MODE` and `GSLIDE_SESSION_PARTITION`.

### Changed
- **Version 1.9.12**, **build 68**.

## [1.9.11] - 2026-04-14

### Added
- **Web UI PIN (optional)** – From desktop **Settings → WAN Access**, set a **4–12 digit** PIN so browsers must unlock at **`/tunnel-unlock`** before the Web UI and its **`/api` proxy** respond. PIN is stored as a **scrypt** hash; unlock uses an **HttpOnly** session cookie (about **7 days**), with **`Secure`** when appropriate (e.g. `trycloudflare.com` or local HTTPS Web UI). Failed attempts are **rate-limited**. Changing or removing the PIN **rotates the session secret** (existing cookies invalidated).
- **PIN scope** – Choose who must unlock: **Cloudflare tunnel only** (default; localhost path while Quick Tunnel is on), **local network only** (non-localhost clients; does not cover the share link, which appears as localhost), or **tunnel and local network**.

### Security / API
- **`GET /api/preferences`** and IPC **`get-preferences`** return a safe payload: `webUiTunnelPinEnabled` and `webUiPinScope` without PIN material. **`POST /api/preferences`** cannot set PIN, scope, tunnel toggle, or controller allowlist (unchanged or tightened for scope).

### Changed
- **Version 1.9.11**, **build 67**.

## [1.9.10] - 2026-04-10

### Added
- **Speaker notes zoom** – Track native Google Slides notes zoom as discrete steps; restore after **Reload presentation** (same pattern as slide index). Preference **Default speaker notes zoom (steps)** (desktop Settings + Web UI) applies when notes open. **`GET /api/status`**: `notesZoomSteps`, `notesZoomDefault`.
- **Companion module** – Variables `notes_zoom_steps` and `notes_zoom_default`; flat **`yarn pack:import`** tarball for Companion import (avoids `npm pack`’s `package/` layout).

### Changed
- **Version 1.9.10**, **build 65**.

## [1.9.9] - 2026-04-07

### Fixed
- **Primary → backup connectivity** – Health checks and command forwarding now resolve the backup HTTP port with **`backupPort` falling back to `apiPort`** (the port the API actually listens on). Saving the API port in **Primary** mode also updates **backup port** (desktop Settings + Web UI) so primaries do not keep probing a stale port. **Verbose logs** record health-check failures (`ECONNREFUSED`, timeouts, non-200). **`GET /api/status`** matches path without query string.
- **Multi-monitor display selection** – Speaker notes now follow the **Speaker notes display** setting. Fullscreen placement uses that display’s bounds instead of inferring from `getDisplayMatching(window.getBounds())` (which could lock notes onto the projector). Saved `notesBounds` are restored only when the window center lies on the selected notes monitor; otherwise notes fullscreen on the correct monitor. `open-presentation` resolves a missing/stale notes display ID with `|| displays[0]` like other code paths. Desktop Settings: persist display IDs as numbers and restore `<select>` values with explicit string coercion.

### Changed
- **Tunnel QR overlay (notes monitor)** – Shows **only** the QR code (no URL caption); **centered** on the notes display; compact window size without the text row.
- **Restricted Web UI (shared / tunnel URL)** – Removed the yellow in-page banner for remote users. The **Settings** tab stays hidden and API proxy restrictions are unchanged. Administrators: see [README.md](README.md) (WAN section), [docs/PUBLIC-ACCESS.md](docs/PUBLIC-ACCESS.md), and [README-SECURITY.md](README-SECURITY.md).

### Build
- **macOS packaging** – `electron-builder` **mac** target now produces **DMG** installers in addition to **ZIP** (arm64 and x64).

## [1.9.8] - 2026-04-06

### Added
- **Cloudflare Quick Tunnel (WAN access)** – Optional Quick Tunnel from desktop Settings; bundles `cloudflared` via `yarn download:cloudflared` and `extraResources`. Uses `--no-tls-verify` when the Web UI origin is HTTPS with a self-signed certificate so the tunnel does not return 502.
- **Safer shared Web UI** – When Quick Tunnel is on, browser sessions that hit the Web UI through the tunnel (localhost via `cloudflared`) see only **Remote** and **Controls**; the **Settings** tab is hidden and selected API routes are blocked at the Web UI proxy. Use the LAN URL for full in-browser settings.
- **Companion tunnel controls** – Four new Companion actions: Enable WAN Tunnel, Disable WAN Tunnel, Show Tunnel QR Code (configurable duration 5–300s), Hide Tunnel QR Code. Two new variables: `tunnel_enabled` (Yes/No) and `tunnel_url` (live URL). New feedback: Tunnel Enabled (blue when active). All tunnel state polled from `/api/status`.
- **Tunnel QR overlay on notes display** – `POST /api/show-tunnel-qr` displays a small 320×360 frameless QR code window in the bottom-right corner of the presenter's notes monitor. Auto-hides after the configured duration; dismissed immediately by `POST /api/hide-tunnel-qr` or when the tunnel stops. Requires `qrcode` package (added).
- **Show QR from web UI Settings** – WAN Access section added to the browser-based Settings tab showing live tunnel status and Show/Hide QR buttons for LAN operators who don't use Companion.
- **Clickable network addresses** – Non-localhost IP addresses in the Electron app's Network Access section now open in the default browser on click.

### Fixed
- **cloudflared download on macOS** – Fetch official `.tgz` archives (bare binary URLs 404).

### Documentation
- README: Quick Tunnel deployment, packaging, and security model; expanded [docs/PUBLIC-ACCESS.md](docs/PUBLIC-ACCESS.md).

## [1.9.7] - 2026-03-18

### Added
- **Speaker notes layout** – Notes Layout setting (Electron + Web UI): hide or narrow the presenter view side panel via CSS injection.
- **Persist speaker notes window bounds** – Notes window size/position saved across app restarts; use `setOpacity(0)` instead of `hide()` to preserve viewport dimensions.

## [1.9.6] - 2026-03-17

### Fixed
- **Backup controls toggle** – Restored app support for decoupling primary/backup via Companion. The Electron app again implements `POST /api/set-backup-controls` and includes `backupControlsEnabled` in `GET /api/status`; primary only forwards commands to backups when enabled. Companion variable and Set Backup Controls action now work; feedback shows "Backups On" / "Backups Off" with distinct styling.

## [1.9.5] - 2026-03-16

### Changed
- **Removed share link / QR overlay feature** – Removed the external redirect + QR overlay system (`/api/share-link`, `/api/show-share-qr`, `/api/hide-share-qr`) that depended on an unavailable external service. Share Settings UI removed from desktop settings.
- **Code cleanup** – Removed task-scaffolding code from Phase 2 development: `backupControlsEnabled` runtime gate, `/api/set-backup-controls` endpoint, CSS injection for speaker notes preview column, and `navigateToSlide()` helper (logic inlined into `/api/go-to-slide`).
- **Repo cleanup** – Removed internal planning docs, test scripts, and draft notes from public repository. Fixed broken `git fetch` caused by macOS `Icon\r` files tracked in git.

## [1.9.3] - 2026-03-15

### Added
- **Backup controls toggle** – Enable/disable backup command forwarding at runtime without restart. New API endpoint: `POST /api/set-backup-controls` `{ "enabled": true|false }`. New Companion action, variable, and feedback.
- **Reload slide position restoration** – Reload now restores slide position via both URL fragment and a post-load `navigateToSlide()` call (1500ms delay for Google Slides JS initialization). Belt-and-suspenders approach for reliability.

### Fixed
- **Speaker notes preview column** – Column now locked to 28% max-width (previously expanded to ~50% on load). CSS injection prevents reflow as preview images stream in.

## [1.9.2] - 2026-02-23

### Added
- **Share link + QR overlay** – Full flow for generating share links and showing a QR code on the presentation display:
  - **Share Settings UI** (desktop) – Share Base URL, Share Register URL, Share API Key; load/save in Settings.
  - **Share link helpers** in main process – `genShareCode()`, `registerShareCode()`, `getShareLink({ forceNew })` with caching (TTL buffer).
  - **API endpoints** – `POST /api/share-link`, `POST /api/show-share-qr`, `POST /api/hide-share-qr` (IP allowlist, JSON).
  - **QR overlay window** – Frameless, transparent window on presentation display; configurable auto-hide (default 20s); uses `qrcode` package.
  - **Companion module actions** – `show_share_qr` (duration 5–300s, force new link) and `hide_share_qr` (27 actions total).
- **Multi-platform builds** – CI and electron-builder now produce:
  - macOS: arm64 (Apple Silicon) and x64 (Intel) .zip.
  - Linux: AppImage (x64 and arm64), and .deb (x64/amd64).
- **Documentation** – RELEASES.md (6 formats, 5 architectures), TESTING.md (all plans), _FINDINGS.md (plan completion), orchestration/token-tracking docs.

### Changed
- **Companion module** – Switched to Yarn (packageManager, .yarnrc.yml, yarn.lock); build uses `yarn install --immutable` and `yarn run package`.
- **Build workflow** – Separate macOS (arm64 + x64) and Linux jobs; release job attaches all artifacts (mac arm64/x64 zips, Linux AppImages + .deb, Companion .tgz).

### Technical
- New npm script: `build:linux` runs all Linux targets from config (AppImage x64/arm64, deb x64).
- macOS build produces both archs; each .app is ad-hoc signed and zipped separately for release.

---

## [1.9.1] - 2025-02-01

### Added
- **Crash reporting and resilience** – Improved error handling and crash reporting.
- **Speaker notes encoding detection** – Detection and handling for speaker notes encoding issues.
- **Apps Script cleanup tool** – Tooling for Apps Script cleanup (see docs/SPEAKER-NOTES-ENCODING.md).

### Changed
- **Web UI** – Styling improvements and preset list updates.
- **Desktop app** – Preset reorder and related desktop UI improvements.

---

## [1.9.0] - 2025-01-22

### Added
- **Speaker notes text normalization** – Notes text is normalized so line breaks display correctly instead of replacement characters (U+FFFD):
  - Server-side: `normalizeSpeakerNotes()` after extraction (replaces `\r\n`, `\r`, `\u2028`, `\u2029` with `\n`; `\uFFFD+` with `\n`; strips `\u0000`).
  - Client-side: same normalization in the Web UI before rendering in the speaker notes panel.
  - Debug: `logFirstReplacementCharContext()` logs index and hex context of the first U+FFFD when verbose logging is on.
- **Scroll Notes buttons** on the Web UI Remote tab – “Scroll Up” and “Scroll Down” to control speaker notes scrolling on the presentation machine (call existing `/api/scroll-notes-up` and `/api/scroll-notes-down`).
- **Reload preserves notes window size** – When speaker notes were open before reload, their window size/position is cached and restored when notes are reopened after reload (`setSpeakerNotesBoundsFromCache`).
- **Open on specific slide on reload** – Reload now opens the presentation URL with a slide fragment (`#slide=id.pN`) so the deck opens on the same slide; the previous go-to-slide round-trip after load was removed.

### Changed
- **Test presentation** – “Open test presentation” now opens the Testpatterns1080p_v3 deck in **presentation mode** directly (uses `/present` URL); no Ctrl+Shift+F5 after load.
- **Test presentation URL** – Updated to `https://docs.google.com/presentation/d/1qKhywpFhjG4tAtA1e2Rk9dB2lVk_uu5_Ol5TaBhvvPo/present`.
- **Presets API** – GET `/api/presets`, POST `/api/presets`, and POST `/api/open-preset` now match on path only (`apiReqPath`), so requests with query strings (e.g. cache-busting) still work.
- **Speaker notes API** – GET `/api/get-speaker-notes` uses path-only matching and responds with `Content-Type: application/json; charset=utf-8` and explicit UTF-8 encoding for the body.
- **Speaker notes window** – When the app opens the speaker notes popup (all flows: test, open by URL, open with notes, reload, open preset), the popup is opened at **full notes-display size** (bounds of the notes monitor) via `getSpeakerNotesWindowOptions(notesDisplay)` to encourage Google Slides’ narrow-preview layout.
- **Speaker notes extraction** – Simplified to a single selector (`div.punch-viewer-speakernotes-text-body-scrollable`); removed the previous long fallback/cleaning logic.
- **toPresentUrl()** – Now accepts an optional slide number and appends `#slide=id.pN` so presentations can be opened on a specific slide (used by reload).

### Fixed
- Speaker notes in the Web UI no longer show broken characters () instead of line breaks; normalization fixes corruption introduced in the pipeline.
- Preset load/save and open-preset from the Web UI and Companion work reliably even when requests include query parameters.

### Technical
- New helpers: `normalizeSpeakerNotes()`, `logFirstReplacementCharContext()`, `getSpeakerNotesWindowOptions()`, `setSpeakerNotesBoundsFromCache()`.
- API request path is normalized once per request: `apiReqPath = String(req.url || '').split('?')[0]` for applicable routes.

---

## [1.8.0] and earlier

See git history and README for features and fixes prior to 1.9.1.
