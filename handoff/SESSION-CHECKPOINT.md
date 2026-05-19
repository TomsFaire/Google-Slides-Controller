# Session Checkpoint — 2026-05-18
*Read this before reading anything else. If it covers current state, skip BUILD-LOG.*

---

## Where We Stopped

Step 2 complete — DMGs built and ready for prod test. macadam 2.0.17 now compiles against Electron 33 (NAPI 10) with source patches. DMGs are at:
- `dist/Google Slides Opener-2.3.2-arm64.dmg` — Apple Silicon
- `dist/Google Slides Opener-2.3.2.dmg` — Intel x64

---

## What Was Decided This Session

- macadam 2.0.18 compiled with SDK 16.0 — incompatible with Desktop Video 14.5 installed on prod → SIGSEGV on any hardware call
- macadam 2.0.17 uses SDK 10.11.2 headers, which Desktop Video 14.5 supports via backwards compatibility
- macadam 2.0.17 source needed 5 patches to compile with Electron 33's NAPI 10:
  1. `capture_promise.cc` line 85: `napi_env` → `node_api_basic_env` for `finalizeCaptureCarrier`
  2. `capture_promise.cc` line 91: `napi_env` → `const napi_env` for `finalizeVideoBuffer` (not called at runtime)
  3. `capture_promise.cc` line 102: `napi_env` → `const napi_env` for `finalizeAudioPacket` (not called at runtime)
  4. `capture_promise.cc` lines 701-702: `napi_create_external_buffer` → `napi_create_buffer_copy` + manual Release()
  5. `capture_promise.cc` lines 857-858: same as above for audio
  6. `playback_promise.cc` line 85: `napi_env` → `node_api_basic_env` for `finalizePlaybackCarrier`
- These patches live in node_modules (not committed) — they are applied at build time by electron-builder's rebuild step, which recompiles from source
- package.json pins macadam to `"2.0.17"` (committed change pending)

---

## Still Open

- User needs to test on prod machine with DeckLink device
- If prod test fails: check dmesg/Console for new crash signal/error
- KG-1: no error handling if `playback.frame()` throws
- KG-2: no `pb.on('error', ...)` handler on macadam playback object

---

## Resume Prompt

Copy and paste this to resume:

---

You are the Architect on this project. Read CLAUDE.md, then ARCHITECT.md.
Confirm where we stopped and what the next action is. Then wait.

---
