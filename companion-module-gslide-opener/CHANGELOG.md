# Companion Module Changelog

## [Unreleased]

### Added
- **Backup controls toggle action** – New action "Set Backup Controls" to enable/disable backup command forwarding at runtime. Includes new variable `backup_controls_enabled` and feedback "Backup Controls Enabled".

### Verified
- **Timer elapsed variable** — Confirmed `timer_elapsed` variable correctly maps from Electron app's `/api/status` endpoint. The app actively scrapes timer values from the presenter view DOM and makes them available to Companion in HH:MM:SS format.

### Stashed for Future Implementation
- **Image preview feedback** — The broken slide image preview feedback has been removed. The underlying `GET /api/get-slide-previews` endpoint remains available in the Electron app for future use.

## [1.4.9] - Previous Release Notes

See git history for earlier versions.
