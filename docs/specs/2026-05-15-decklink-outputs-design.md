# DeckLink Outputs — Design Spec

**Date:** 2026-05-15  
**Status:** Implemented on `feature/decklink-output`

---

## Problem

Live production events need the app's slides and notes content as clean professional video signals feeding hardware vision mixers, recording decks, and SDI/HDMI displays. The existing Electron windowing model only routes output to OS-visible monitors; there is no path to DeckLink hardware.

## Solution

Capture the slides and notes BrowserWindows frame-by-frame using Electron's `webContents.capturePage()` and push the BGRA buffers to a DeckLink output device. Two independent providers are supported:

- **Primary — macadam** ([Streampunk/macadam](https://github.com/Streampunk/macadam)): NAPI native Node addon for the Blackmagic DeckLink SDK. Low latency, hardware-clock paced.
- **Fallback — FFmpeg**: Pipe raw BGRA frames to an FFmpeg child process using `-f decklink`. Requires FFmpeg built with DeckLink support.

If neither is available the app starts normally with DeckLink outputs marked unavailable.

---

## Architecture

```
DecklinkOutputManager (singleton, src/decklink-output.js)
│
├── detectProviderType()   — macadam → ffmpeg → unavailable (cached)
│
├── SlidesProvider (MacadamProvider or FfmpegProvider)
│   └── OutputController   — capturePage() interval → resize → pushFrame()
│
└── NotesProvider  (MacadamProvider or FfmpegProvider)
    └── OutputController   — capturePage() interval → resize → pushFrame()
```

Each `OutputController` runs a `setInterval` loop at the configured framerate. On each tick:
1. Call `webContents.capturePage()` on the target window
2. Resize the NativeImage to the configured DeckLink output resolution
3. Convert to a BGRA buffer via `toBitmap()`
4. Push to the provider

If the window is absent or destroyed, a black frame is pushed. If `capturePage()` fails, the last good frame is held for up to 2 seconds then replaced with black.

---

## Configuration

Stored in `userData/preferences.json` under a `decklink` key:

```json
"decklink": {
  "slides": { "enabled": false, "deviceIndex": 0, "displayMode": "1080p2997" },
  "notes":  { "enabled": false, "deviceIndex": 1, "displayMode": "1080p2997" }
}
```

Supported display modes: `1080p2997`, `1080p25`, `1080p30`, `1080i5994`, `720p5994`, `720p50`.

---

## Integration Points

| Location | Change |
|----------|--------|
| `src/decklink-output.js` | New module — all DeckLink/FFmpeg logic |
| `main.js:30` | Guarded `require('./src/decklink-output')` |
| `main.js:~1258` | `decklink` defaults in `loadPreferences()` |
| `main.js:~3271` | IPC handlers: `get-decklink-devices`, `get-decklink-status`, `save-decklink-config` |
| `main.js:~6215` | HTTP endpoints: `GET /api/decklink/status`, `GET /api/decklink/devices` |
| `main.js:whenReady` | `DecklinkOutputManager.init()` after `startHttpServer()` |
| `main.js:before-quit` | `DecklinkOutputManager.shutdown()` |
| `preload.js` | Expose `getDecklinkDevices`, `getDecklinkStatus`, `saveDecklinkConfig` |
| `index.html` | DeckLink tab in sidebar + section with device/format selects and status LEDs |
| `renderer.js` | DOM bindings, `populateDecklinkDevices`, `loadDecklinkSettings`, `updateDecklinkStatus`, save handler, 5s poll on tab focus |
| `package.json` | `macadam` as `optionalDependency`; `@electron/rebuild` as `devDependency`; `rebuild` script |

---

## Build Notes

`macadam` is a native Node addon and must be compiled against Electron's ABI headers (not stock Node.js headers):

```bash
yarn rebuild        # runs electron-rebuild -f -w macadam
yarn build:mac      # must run after yarn rebuild on each machine
```

The Blackmagic DeckLink drivers and desktop video SDK must be installed at runtime. The app degrades gracefully if they are absent.

---

## HTTP API

| Method | Endpoint | Purpose |
|--------|----------|---------|
| GET | `/api/decklink/status` | Active provider, per-output active state |
| GET | `/api/decklink/devices` | Enumerated DeckLink device names and indices |

Both endpoints are read-only and subject to the IP allowlist.

---

## Verification Checklist

- [ ] App starts without DeckLink hardware — both outputs show "unavailable"
- [ ] `GET /api/decklink/status` returns correct shape with no hardware
- [ ] Enable slides output in settings → slides content appears on DeckLink device 0
- [ ] Enable notes output in settings → notes content appears on DeckLink device 1
- [ ] Close presentation window → black frame on slides DeckLink output
- [ ] FFmpeg fallback: rename macadam `.node` file, restart → provider shows `ffmpeg`
- [ ] Format change → output reconfigures to new resolution/framerate
- [ ] `yarn build:mac` → macadam rebuilds against Electron headers, included in `.app`
