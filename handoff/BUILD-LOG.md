# Build Log
*Owned by Architect. Updated by Builder after each step.*

---

## Current Status

**Active step:** Step 2 — DeckLink SDK version compatibility fix (macadam 2.0.17 + DV 14.5)
**Last cleared:** none
**Pending deploy:** YES — DMGs built, awaiting prod test

---

## Step History

### Step 1 — Fix MacadamProvider (worker exits bug) — COMPLETE
*Date: 2026-05-18*

Root cause: uncommitted `decklink-output.js` moved `macadam.playback()` into a `child_process.fork()` worker. DeckLink SDK requires main process context; 8MB/frame IPC also untenable. Fix: revert `MacadamProvider` to main process; keep worker for probe + get_devices only.

Files to change:
- `src/decklink-output.js` — revert MacadamProvider class to main-process approach
- `src/macadam-worker.js` — remove start/frame/stop commands (new file, untracked)

Decisions made:
- Keep `workerRpc()` and `macadam-worker.js` for probe + device enumeration (good design, keep it)
- Do NOT use worker for live frame pushing (performance + DeckLink SDK incompatibility)
- Keep `_cachedProviderType = null` reset in `reconfigure()` (correct bug fix from uncommitted work)

Reviewer findings: no must-fix; crash isolation trade-off confirmed intentional
Deploy: committed a8bde76, pushed feature/decklink-output 2026-05-18

---

### Step 2 — macadam 2.0.17 NAPI 10 source patch (DV 14.5 compatibility) — COMPLETE (build only)
*Date: 2026-05-18*

Root cause: macadam 2.0.18 was compiled against DeckLink SDK 16.0; prod machine has Desktop Video 14.5. Any DeckLink hardware SDK call → SIGSEGV. macadam 2.0.17 uses SDK 10.11.2 headers, compatible with DV 14.5 at runtime.

Problem: macadam 2.0.17 source cannot compile against Electron 33 (NAPI 10) without patches:
1. `napi_create_external_buffer` — removed in NAPI 10; replaced with `napi_create_buffer_copy` in `capture_promise.cc` (lines 701, 857). Frames are copied instead of zero-copy; acceptable since we only use playback not capture.
2. `napi_finalize` callback `const napi_env` → `node_api_basic_env` required by `napi_create_external` in NAPI 10; patched in `finalizeCaptureCarrier` (capture_promise.cc:85) and `finalizePlaybackCarrier` (playback_promise.cc:85).

Files patched (in node_modules, NOT committed — patch is build-time only):
- `node_modules/macadam/src/capture_promise.cc` — 4 changes
- `node_modules/macadam/src/playback_promise.cc` — 1 change

package.json change: macadam pinned to `"2.0.17"` in `optionalDependencies` (was `"^2.0.0"`).

DMGs built:
- `dist/Google Slides Opener-2.3.2-arm64.dmg` (120 MB) — for Apple Silicon prod machine
- `dist/Google Slides Opener-2.3.2.dmg` (124 MB) — x64 (Intel)

Deploy: PENDING — user to test on prod machine with DeckLink device

---

## Known Gaps
*Logged here instead of fixed. Addressed in a future step.*

- **KG-1** — `MacadamProvider.pushFrame()` has no error handling if `playback.frame()` throws — logged 2026-05-18
- **KG-2** — No `pb.on('error', ...)` handler on macadam playback object — logged 2026-05-18

---

## Architecture Decisions
*Locked decisions that cannot be changed without breaking the system.*

- Worker process (`macadam-worker.js`) scope is probe + device enumeration only. Live frame pushing stays in main process. — 2026-05-18
- NAPI 10 patches to macadam C++ source (6 changes across `capture_promise.cc` and `playback_promise.cc`) are now reproduced automatically via `scripts/patch-macadam.js`, which runs as a `postinstall` hook after every `yarn install`. The script is idempotent and exits 0 gracefully when macadam is absent (optional dep). Raw patch details are in `handoff/SESSION-CHECKPOINT.md`. — 2026-05-20
