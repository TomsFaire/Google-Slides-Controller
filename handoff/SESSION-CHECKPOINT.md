# Session Checkpoint — 2026-05-18
*Read this before reading anything else. If it covers current state, skip BUILD-LOG.*

---

## Where We Stopped

Pending DV update on prod machine. DMGs at:
- `dist/Google Slides Opener-2.3.2-arm64.dmg` — Apple Silicon (macadam 2.0.18, SDK 10.11.2)
- `dist/Google Slides Opener-2.3.2.dmg` — Intel x64

User is updating Desktop Video to latest stable. After that: install the arm64 DMG and retest.

---

## What Was Decided This Session

- ALL macadam versions (2.0.14–2.0.18, the entire npm history) use DeckLink SDK 10.11.2 — no newer SDK available
- The crash is INSIDE DV 14.5's DeckLinkAPI.framework, not in macadam's wrapper. SDK version on our side is irrelevant.
- Root cause: DV 14.5 is incompatible with the macOS version running on prod (likely macOS 15 Sequoia KEXT issues, or a DV 14.5 regression bug)
- Decision: Update Desktop Video to latest stable; macadam 2.0.18 source compiled from scratch for Electron 33 with NAPI 10 patches
- macadam source (same in 2.0.17 and 2.0.18) requires 5 NAPI 10 patches applied at build time:
  1. `capture_promise.cc`: `finalizeCaptureCarrier` → `node_api_basic_env` (napi_create_external callback)
  2. `capture_promise.cc`: `finalizeVideoBuffer` → `const napi_env` (no longer a callback)
  3. `capture_promise.cc`: `finalizeAudioPacket` → `const napi_env` (no longer a callback)
  4. `capture_promise.cc`: `napi_create_external_buffer` → `napi_create_buffer_copy` + immediate Release() (video)
  5. `capture_promise.cc`: same as above for audio packet
  6. `playback_promise.cc`: `finalizePlaybackCarrier` → `node_api_basic_env` (napi_create_external callback)

Before installing DMG after DV update — also check:
- System Settings → Privacy & Security — approve Blackmagic Design kernel extension if prompted
- DeckLink device shows in Desktop Video Setup app after update

---

## Still Open

- Prod test after DV update
- KG-1: no error handling if `playback.frame()` throws
- KG-2: no `pb.on('error', ...)` handler on macadam playback object

---

## Resume Prompt

Copy and paste this to resume:

---

You are the Architect on this project. Read CLAUDE.md, then ARCHITECT.md.
Confirm where we stopped and what the next action is. Then wait.

---
