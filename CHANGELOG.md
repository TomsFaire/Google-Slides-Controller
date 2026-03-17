# Changelog

All notable changes to Google Slides Opener are documented here.

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
