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

## Status — review fixes (PR #19, still draft)

`feature/offline-mode` has been merged up to `main` (post-v2.4.0, DeckLink output) and the
code-review findings 1–7 are addressed. Conflicts were resolved keeping both sides; the only
deliberate drops were two places where `main` had moved on independently of this feature
(stage-timer-overlay enabled state now reads the preference, and `presentationGpuMode` stays a
frozen `'default'` literal in `/api/status`).

Decisions made while fixing:

- **Cache states** are now a frozen enum in `src/offline-cache.js` and include `'error'`. A warm
  that fails lands there instead of sitting on `'caching'`; both the desktop panel and the web
  remote render it as a red "Failed" badge.
- **Warm-ups are serialised** on the in-flight promise (`createWarmGuard`), not on the cache
  state, because the state is set inside the async run and overlapping callers could both pass a
  state check. A second request is **refused**, not queued: IPC returns
  `{ success: false, error: 'Cache warm-up already in progress' }` and
  `POST /api/offline/warm-cache` answers **409**. This is an API contract change from the
  original draft, which always answered 200.
- **Presenter position** is captured before the warm starts and restored in `finally`, so the
  success, error and early-return paths all put the deck back. The end slide is read live from
  the notes window rather than counted from key presses, since a Right arrow that only advances a
  click-build does not change slide. Restore key events are paced at 100 ms like
  `/api/go-to-slide`.
- **Manifest reads** go through `readManifestEntries()`, which returns `null` for a corrupted or
  version-mismatched `cache-manifest.json` and drops individual entries with no usable `hash`.

Pending (needs hardware / a signed-in Google session, cannot be verified in CI):

- Auto-warm leaving the presenter on their original slide on a real deck
- Offline playback itself (enable → pull the network → navigate)
- The 409 path as seen by the web remote and Companion
- Whether 100 ms restore pacing is enough for large decks
