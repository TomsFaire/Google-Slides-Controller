# Offline Mode Research Findings

## How the app currently opens presentations

- **URL:** `https://docs.google.com/presentation/d/{id}/present` (no query params)
- **Window type:** `BrowserWindow` (not webview)
- **Session partition:** `persist:google` — persistent across app restarts
- **webPreferences:** `nodeIntegration: false`, `contextIsolation: true`
- **URL builder:** `toPresentUrl()` in `main.js:469`

## Hard constraint discovered

Google Slides `/present` mode **cannot function offline.** This is enforced by Google's own client-side JavaScript — the presentation shell detects network absence and refuses to render slides. This is confirmed by Google Support: "Presentation mode is not available while offline."

This applies even when the Google Docs Offline Chrome extension is installed. The extension enables offline editing of documents, not offline presentation of slides.

## URL parameters tested (research only)

| Parameter | Result |
|-----------|--------|
| `?offline=true` | Not a supported parameter — no effect |
| `?start=true` | Supported: auto-starts the presentation |
| `?loop=true` | Supported: loops the presentation |
| `?delayms=N` | Supported: auto-advances slides |

No URL parameter triggers aggressive asset caching or unlocks offline presentation mode.

## Service Worker

- **Present:** NO — Service Worker only registers when the **Google Docs Offline Chrome extension** is installed and activated in Drive Settings
- **Cache names found:** N/A (extension not installed in this Electron context)
- **Without extension:** Standard HTTP caching only (CSS/JS cached normally by Chromium)

## Assets that cache automatically (via HTTP caching in `persist:google`)

| Asset type | HTTP-cached | Survives offline |
|------------|-------------|-----------------|
| Slide metadata/JSON | Partial | NO (Google blocks offline) |
| Google Drive images | YES | NO (blocked by client-side check) |
| External URL images | YES (if cached) | NO (blocked by client-side check) |
| Google Fonts | YES | NO (blocked by client-side check) |
| Google Drive videos | NO | NO |
| YouTube embeds | NO | NO |

Even when assets are in the HTTP cache, the presentation shell itself refuses to function when Google's servers are unreachable.

## Electron-specific notes

- Persistent partition means HTTP cache, cookies, and IndexedDB persist across restarts — but this is insufficient for offline operation
- Electron *can* load Chrome extensions via `session.loadExtension()` — loading the Google Docs Offline extension into `persist:google` is a potential path but:
  - Requires the unpacked extension (not from Chrome Web Store directly)
  - Extension would need activation via Drive Settings UI
  - Still would not unlock `/present` offline per Google's own statement

## What `offlineModeEnabled` does in v2.5.0

Given the above constraints, the v2.5.0 implementation builds the plumbing:
- Preference key, IPC handlers, API routes, desktop UI toggle, web remote toggle
- When enabled: marks the presentation as "cache warmed" after it has been open for a configurable delay
- Does **not** isolate the network or guarantee offline operation

## Day 2 items

1. **Chrome extension injection** — load Google Docs Offline extension into `persist:google` session via `session.loadExtension()`, then auto-enable it via Drive Settings automation
2. **PDF fallback rendering** — export current presentation via Slides API, render PDF locally when network drops
3. **Asset audit report** — scan open presentation DOM for external asset URLs, warn operator which won't survive offline
4. **Manual pre-cache** — "Download for offline" button that crawls slide image URLs and forces them into HTTP cache
5. **Partial network isolation** — block non-Google CDN requests after warm-up to prevent external asset timeouts from affecting slides

## Version target

Infrastructure (toggle, API, UI) ships in v2.5.0. Functional offline rendering deferred to v2.6.0 pending Day 2 research.
