# Companion Module Changelog

## [Unreleased]

### Verified
- **Timer elapsed variable** — Confirmed `timer_elapsed` variable correctly maps from Electron app's `/api/status` endpoint. The app actively scrapes timer values from the presenter view DOM and makes them available to Companion in HH:MM:SS format.

### Stashed for Future Implementation
- **Image preview feedback** — The broken slide image preview feedback has been removed. The underlying `GET /api/get-slide-previews` endpoint remains available in the Electron app for future use.

## [1.4.9] - Previous Release Notes

See git history for earlier versions.
