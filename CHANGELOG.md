# Changelog

All notable changes to Google Slides Opener are documented here.

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

See git history and README for features and fixes prior to 1.9.0.
