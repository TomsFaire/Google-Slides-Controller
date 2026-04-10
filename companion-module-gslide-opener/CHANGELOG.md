# Companion Module Changelog

## [1.4.9] - 2026-04-10

### Added
- **Speaker notes zoom variables** – `notes_zoom_steps` (live offset from Slides baseline) and `notes_zoom_default` (saved preference default), mapped from Electron app `GET /api/status` (`notesZoomSteps`, `notesZoomDefault`). See companion HELP.

## [1.4.8] - 2026-04-07

### Fixed
- **Tunnel QR actions** – Replaced removed endpoints `/api/show-share-qr` and `/api/hide-share-qr` with **`/api/show-tunnel-qr`** and **`/api/hide-tunnel-qr`**. Request body uses **`duration`** (seconds) as required by the app. Resolves Companion error `Failed to show QR: Not found` (HTTP 404). Action IDs `show_share_qr` / `hide_share_qr` are unchanged so existing buttons keep working; labels now say **Tunnel QR**. Removed obsolete "Generate New Share Link" option.

## [1.4.7] - 2026-04-07

### Changed
- **Version bump** – Aligns packaged module metadata with `package.json`; `companion/manifest.json` version synced to match.

## [1.4.6] - 2026-03-15

### Added
- **Backup controls toggle action** – New action "Set Backup Controls" to enable/disable backup command forwarding at runtime. Includes new variable `backup_controls_enabled` and feedback "Backup Controls Enabled".

### Verified
- **Timer elapsed variable** — Confirmed `timer_elapsed` variable correctly maps from Electron app's `/api/status` endpoint. The app actively scrapes timer values from the presenter view DOM and makes them available to Companion in HH:MM:SS format.

### Stashed for Future Implementation
- **Image preview feedback** — The broken slide image preview feedback has been removed. The underlying `GET /api/get-slide-previews` endpoint remains available in the Electron app for future use.

## [1.4.5] - Previous Release Notes

See git history for earlier versions.
