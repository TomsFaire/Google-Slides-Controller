# Native DeckLink Addon (replace macadam) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the unmaintained `macadam` dependency with a small in-repo, output-only native addon built against current Blackmagic SDK headers, and use the move to remove the per-frame copy and CPU-conversion costs.

**Architecture:** A Node-API (stable-ABI) C++ addon, loaded only inside the existing forked worker (`src/macadam-worker.js`, to be renamed), owns the DeckLink device, converts BGRA to UYVY natively, and runs its own output pacing (a "latest frame wins" slot that is repeated at the output rate). A second tiny addon module (`shm`) lets the Electron main process write captured frames into shared memory so the worker never receives frames over IPC. A backend adapter seam lets `macadam` and `native` run side by side on the feature branch for hardware A/B only; once the native backend passes the parity checklist the default switches and macadam is deleted in the same release (no release ships both).

**Tech Stack:** C++17, node-addon-api (Node-API pinned to a fixed `NAPI_VERSION`), node-gyp via `electron-rebuild`, Blackmagic Desktop Video SDK 16.x headers (vendored), Apple Accelerate/vImage for conversion, `node:test` for unit tests.

**Spec:** Requirements below (from the 2026-10-20 design discussion in PR #23) plus the original [`docs/specs/2026-05-15-decklink-outputs-design.md`](../specs/2026-05-15-decklink-outputs-design.md). Nothing in this plan starts until PR #23's hardware checklist is complete and merged.

## Requirements

- **R1 Maintainability.** No dependency on an unmaintained package. SDK headers are ours to update. A new Electron major must not require patching native sources.
- **R2 Frame rate.** 1080p at 24 fps minimum, 30 fps preferred, 60 fps nice to have, measured end to end with video/animation content on a 1080p 1x display. Output signal stays at the nominal rate even when capture is slower.
- **R3 Compatibility.** Any DeckLink/UltraStudio device with video output, at modes up to 1080p60 (UHD/keying are out of scope for this plan, see Task 8; decided order: get 1080p solid without macadam first, then UHD). Capture-only devices are not listed. Pixel format (BGRA vs 8-bit YUV) is negotiated with the device, not hard-coded per mode.
- **R4 Safety.** Native crashes stay in the worker. A missing/old driver or an unsupported mode produces a clear error in the DeckLink tab, never a crash or a hang.
- **R5 Diagnostics.** The DeckLink tab and logs report the installed DeckLink API version and per-stage timings.
- **R6 Parity before removal.** macadam is deleted only after the native backend matches every item in Task 7's checklist on real hardware. The switch is outright: no release keeps macadam as a runtime fallback.

## Global Constraints

- Platform for this plan: macOS 13+, arm64 and x64. Windows/Linux are not targeted.
- Hardware tests cannot run in the Claude sandbox and `yarn install` / native builds are broken there: steps marked **[user]** are run by the user in their own terminal with the UltraStudio attached.
- Hardware tests run under Electron's Node: `ELECTRON_RUN_AS_NODE=1 electron …` (the addon is built for Electron).
- Version bump only the patch digit unless told otherwise, and never on a feature branch; increment `buildNumber` before any packaging build.
- Never commit `package.json` dependency changes without a regenerated `yarn.lock` (user regenerates it).
- The addon links `DeckLinkAPI.framework` at runtime via the SDK's dispatch loader (no link-time dependency), so the app must still start with Desktop Video absent.
- Commit trailer: `Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>`.

## Review Focus

Failure modes the requirements imply that no single task's happy path covers. Each has a test in the task that owns the code.

1. **Capture stalls but the output must not go black or freeze the device** (a slow `capturePage()`): the output repeats the last frame at the nominal rate. Test in Task 4.
2. **Driver missing, or older/newer than the addon understands:** the addon loads, `version()` reports, and `listDevices()` returns an error object rather than throwing at require time. Test in Task 2.
3. **Mode or pixel format the device rejects:** the error lists the modes the device does support, and BGRA to YUV fallback is tried before failing. Test in Task 3.
4. **Device unplugged or taken by another app mid-show:** the worker reports a runtime error, the app keeps running, the status LED shows the error, and a later reconfigure succeeds. Test in Task 4.
5. **Worker dies mid-show:** the app survives and the controller stops pushing (no write to a dead pipe, no unhandled `error`). Test in Task 1.

---

## File Structure

| Path | Responsibility |
|---|---|
| `src/decklink-backend.js` (new) | Backend interface + `createBackend(name)`; the only place that knows macadam vs native |
| `src/backends/macadam-backend.js` (new) | The current macadam calls, moved out of the worker |
| `src/backends/native-backend.js` (new) | Wraps the native addon behind the same interface |
| `src/decklink-modes.js` (new) | `DISPLAY_MODES` and the mapping from our mode keys to BMD mode names (moved out of `decklink-output.js`) |
| `src/decklink-worker.js` (renamed from `macadam-worker.js`) | Worker message loop; calls the backend only |
| `native/decklink/` (new) | `binding.gyp`, addon sources, vendored SDK headers (`sdk/`, version recorded in `sdk/VERSION`) |
| `native/shm/` (new) | Tiny shared-memory addon, safe to load in the Electron main process |
| `scripts/decklink-native-info.js` (new) | Hardware script: version, devices, modes, pixel formats |
| `tests/decklink-backend.test.js` (new) | Backend/worker logic against a fake backend |
| `tests/native-parity.test.js` (new) | JS reference vs native conversion; skipped when the addon is not built |

---

### Task 1: Backend seam (no behaviour change)

**Files:**
- Create: `src/decklink-backend.js`, `src/backends/macadam-backend.js`, `tests/decklink-backend.test.js`
- Modify: `src/macadam-worker.js` (calls the backend only), `src/decklink-output.js` (`MacadamProvider` becomes `WorkerProvider`; `WORKER_SCRIPT` constant)

**Interfaces:**
- Produces: `createBackend(name: 'macadam' | 'native'): Backend` where
  `Backend = { probe(): Promise<{ driverVersion: string | null }>, listDevices(): Promise<Array<{ index: number, name: string, canOutput: boolean }>>, openOutput(opts: { deviceIndex: number, mode: string, width: number, height: number, pixelFormat: 'bgra' | 'yuv' }): Promise<Output>, }` and
  `Output = { displayFrame(bgra: Buffer): Promise<{ convertMs: number, displayMs: number }>, close(): Promise<void> }`.
  Backend name comes from `process.env.DECKLINK_BACKEND`, default `'macadam'`.

- [ ] **Step 1: Write the failing test** in `tests/decklink-backend.test.js` using an injected fake backend: `worker handle('frame')` returns the timing object from `displayFrame`; `handle('start')` rejects with the backend's error message when `openOutput` throws; a second `frame` while the first is pending replies `skip`.
- [ ] **Step 2: Run** `yarn test` — expect FAIL (`createBackend` not defined).
- [ ] **Step 3: Implement** `createBackend` and move the macadam calls (probe, `getDeviceInfo`, `playback`, `displayFrame`, `stop`, YUV conversion via `bgraToUyvy`) from the worker into `macadam-backend.js` unchanged. The worker keeps only message routing, the busy flag and the 2000 ms `displayFrame` timeout. Export a `handle(msg, backend)` function for tests.
- [ ] **Step 4: Run** `yarn test` — expect all tests pass (136 + new).
- [ ] **Step 5: Test for Review Focus 5** — `WorkerProvider` after the fake worker emits `exit`: `pushFrame` returns `false` and no write is attempted. Make it pass.
- [ ] **Step 6: [user] Hardware regression:** `yarn test:decklink-hw` at 1080p60, 1080p59.94, 1080p29.97 — 7/7 each. This proves the refactor changed nothing.
- [ ] **Step 7: Commit** `refactor(decklink): put macadam behind a backend interface`

### Task 2: Native addon skeleton and build integration

**Files:**
- Create: `native/decklink/binding.gyp`, `native/decklink/src/addon.cc`, `native/decklink/sdk/*` (headers + `DeckLinkAPIDispatch.cpp` from SDK 16.x, `VERSION`), `scripts/decklink-native-info.js`
- Modify: `package.json` (scripts `build:native`, `decklink:native-info`; `asarUnpack` for the built `.node`), `.github/workflows/build.yml` (build the addon before packaging)

**Interfaces:**
- Produces: `require('native/decklink/build/Release/decklink.node')` exporting `version(): { api: string, sdk: string }` and `listDevices(): { devices: Array<{...}>, error: string | null }` (empty devices plus an error string when the driver is missing; never throws at require time).

- [ ] **Step 1: Vendor the SDK headers** (decided 2026-10-20: yes). Copy the SDK licence text verbatim into `native/decklink/sdk/LICENSE` and record the SDK version in `native/decklink/sdk/VERSION`. macadam already ships SDK 10.11.2 headers under the same licence text.
- [ ] **Step 2: Implement** the addon with `node-addon-api`, `NAPI_VERSION` pinned (state the chosen value in `binding.gyp`), compiling the SDK's own dispatch file, so the addon links no Blackmagic library at build time.
- [ ] **Step 3: Test** `tests/native-smoke.test.js` (skips when the `.node` is absent): `version().api` matches `/^\d+\.\d+/`; `listDevices()` returns an object with `devices` array and `error` either `null` or a string (Review Focus 2).
- [ ] **Step 4: [user]** `yarn build:native` then `ELECTRON_RUN_AS_NODE=1 electron scripts/decklink-native-info.js` — prints the API version matching the macadam `deckLinkVersion()` and the UltraStudio by name.
- [ ] **Step 5: CI** — `build-macos` builds the addon for arm64 and x64 and the packaged app contains the `.node` unpacked. Green run is the check.
- [ ] **Step 6: Commit** `feat(decklink): native addon skeleton with SDK 16 headers`

### Task 3: Device and mode enumeration with pixel-format negotiation

**Files:**
- Create: `src/decklink-modes.js`
- Modify: `native/decklink/src/addon.cc` (add `listDevices` fields, `supportedModes`), `src/backends/native-backend.js` (new), `src/decklink-backend.js`, `src/decklink-output.js` (import modes from `decklink-modes.js`; drop the static `pixelFormat: 'yuv'` flags)

**Interfaces:**
- Produces: `supportedModes(deviceIndex): Array<{ mode: string, width: number, height: number, fps: number, pixelFormats: Array<'bgra' | 'yuv'> }>` from the addon (uses `DoesSupportVideoMode` for each pixel format). `choosePixelFormat(supported: Array<'bgra'|'yuv'>): 'bgra' | 'yuv'` in `decklink-modes.js` prefers `bgra` (no conversion).
- `listDevices()` entries gain `canOutput: boolean` (false for recorder/capture-only devices).

- [ ] **Step 1: Failing tests** in `tests/decklink-backend.test.js`: `choosePixelFormat(['yuv']) === 'yuv'`, `choosePixelFormat(['bgra','yuv']) === 'bgra'`, `choosePixelFormat([])` throws an error whose message lists the device's supported modes (Review Focus 3); the output-device filter drops `canOutput: false` entries.
- [ ] **Step 2: Run** — expect FAIL. **Step 3: Implement** `decklink-modes.js` (move `DISPLAY_MODES`, add the helpers) and make the DeckLink tab device list use the filtered list.
- [ ] **Step 4: Run** `yarn test` — expect pass.
- [ ] **Step 5: Implement the native `supportedModes`/`canOutput`** and the native backend's `listDevices`.
- [ ] **Step 6: [user]** `decklink-native-info.js` on the UltraStudio 3G reports 1080p50/59.94/60 as `['yuv']` and 1080p29.97/30/25 as `['bgra','yuv']`, matching what `decklink:probe-modes` found earlier.
- [ ] **Step 7: Commit** `feat(decklink): negotiate pixel format with the device`

### Task 4: Native output with its own pacing

**Files:**
- Modify: `native/decklink/src/addon.cc` (+ `output.cc`/`output.h` if the file passes ~600 lines), `src/backends/native-backend.js`, `src/decklink-worker.js`, `src/decklink-output.js` (OutputController pushes a frame and stops caring about busy/drops)

**Interfaces:**
- Produces: `openOutput({deviceIndex, mode, pixelFormat}): Output` with `setFrame(bgra: Buffer): void` (non-blocking; replaces the slot), `stats(): { scheduled: number, late: number, dropped: number, buffered: number }`, `close(): Promise<void>`. A native thread schedules frames with `ScheduleVideoFrame` and a small preroll queue, re-sending the newest slot contents at the output rate. A `'error'` callback fires on device loss.

- [ ] **Step 1: Failing JS test** with a fake native `Output`: when `setFrame` is not called for 2 seconds the controller still reports healthy and does not push black (Review Focus 1); a fake `'error'` event sets the provider status to an error string and a later `reconfigure()` opens a fresh output (Review Focus 4).
- [ ] **Step 2–4:** implement, run, pass.
- [ ] **Step 5: [user] Hardware:** `DECKLINK_BACKEND=native yarn test:decklink-hw` at 1080p60, 59.94, 29.97 — 7/7. Unplug the USB4 cable mid-run: the app stays up and the tab shows an error; plugging it back in and saving settings restores output.
- [ ] **Step 6: Commit** `feat(decklink): native scheduled output that holds the last frame`

### Task 5: Native BGRA→UYVY conversion

**Files:**
- Modify: `native/decklink/src/addon.cc` (conversion inside `setFrame` using vImage), `src/backends/native-backend.js`
- Test: `tests/native-parity.test.js`

**Interfaces:**
- Consumes: `bgraToUyvy(bgra, w, h, out, { matrix, range })` from `src/bgra-to-uyvy.js` as the reference.
- Produces: `openOutput` option `{ matrix: '709' | '601', range: 'limited' | 'full' }` (defaults 709/limited, same env overrides as the worker today).

- [ ] **Step 1: Failing test** (skips without the addon): for a 64×32 frame of known patches (the colour-bar values used in `tests/bgra-to-uyvy.test.js`) the native result equals the JS reference within ±1 per byte, for each of 709/601 × limited/full.
- [ ] **Step 2–4:** implement with vImage, run, pass.
- [ ] **Step 5: [user]** `yarn decklink:color-chart` through the native backend looks identical to the macadam output on the Video Assist.
- [ ] **Step 6: Commit** `feat(decklink): native colour conversion matching the JS reference`

### Task 6: Shared-memory frame transfer

**Files:**
- Create: `native/shm/binding.gyp`, `native/shm/src/shm.cc`, `tests/shm.test.js`
- Modify: `src/decklink-output.js` (write the captured bitmap into the slot, send `{slot}` only), `src/decklink-worker.js`, `native/decklink/src/addon.cc` (`setFrameFromShm(name, slot)`)

**Interfaces:**
- Produces: `shm.create(name: string, slotBytes: number, slots: number): Region`, `shm.open(name: string): Region`, `Region.slot(i: number): Buffer` (a view, no copy), `Region.close(): void`. Two slots minimum (write one while the worker reads the other).

- [ ] **Step 1: Failing test** `tests/shm.test.js` (skips without the addon): a forked child opens the region, reads the bytes the parent wrote to slot 0, and writes a different pattern to slot 1 that the parent then reads.
- [ ] **Step 2–4:** implement with `shm_open`/`mmap`, run, pass.
- [ ] **Step 5:** wire it in: the main process stays free of any DeckLink SDK calls (only the shm module is loaded there).
- [ ] **Step 6: [user]** compare the stats lines before and after: worker round trip drops from the macadam baseline (to be measured in PR #23's tests) to well under the frame interval, and `frames dropped` is 0 at 1080p30.
- [ ] **Step 7: Commit** `perf(decklink): hand frames to the worker through shared memory`

### Task 7: Parity gate, switch the default, remove macadam

**Files:**
- Modify: `package.json` (remove `macadam` dependency, `rebuild` script, `asarUnpack` entry), `.github/workflows/build.yml`, `README.md`, `docs/DECKLINK.md`
- Delete: `src/backends/macadam-backend.js`, `scripts/patch-macadam.js`, the `postinstall` hook; user regenerates `yarn.lock`

**Parity checklist (all **[user]**, all on the native backend, all must pass before any deletion):**

- [ ] `yarn test:decklink-hw` 7/7 at 1080p60, 59.94, 29.97, plus a clean `yarn decklink:probe-modes` over every mode in the app
- [ ] Colour chart identical to macadam's output on the Video Assist
- [ ] Real deck on a 1080p 1x display: sustained ≥ 24 fps with a video slide, stats lines recorded
- [ ] Slide changes follow within ~0.5 s for a 3-minute run (no latency creep)
- [ ] Packaged `.app` (not `yarn start`) outputs on the hardware
- [ ] Desktop Video absent: the app starts, the tab explains why
- [ ] Recorded: the Desktop Video / firmware versions tested

- [ ] **Step 1:** once the checklist passes, in one change flip nothing to a flag: make `native` the only backend, delete `src/backends/macadam-backend.js`, `scripts/patch-macadam.js`, the `postinstall` hook, the `macadam` dependency, its `rebuild` script and `asarUnpack` entry; remove `DECKLINK_BACKEND`; the user regenerates `yarn.lock`.
- [ ] **Step 2:** update `README.md` and `docs/DECKLINK.md`: drop macadam mentions, record the tested Desktop Video / firmware versions.
- [ ] **Step 3:** `yarn test` green, `build-macos` green, packaged `.app` re-checked on hardware.
- [ ] **Step 4: Commit** `chore(decklink): replace macadam with the native addon`

### Task 8 (separate plan, starts only after Task 7 ships)

UHD output and external keying (fill + key from one device) build on this addon; each needs its own spec and a key/fill-capable device to test on. Open items to carry over: premultiplied vs straight alpha, BGRA vs YUVA on keying devices, sub-device profiles on Duo 2/Quad 2.

## Self-Review notes

- **Coverage:** R1 → Tasks 2, 7; R2 → Tasks 4–6 and the gate; R3 → Task 3; R4 → Tasks 2, 4 and Review Focus; R5 → Task 2 (version), 4 (`stats()`); R6 → Task 7.
- **Decisions (2026-10-20):** vendor the SDK headers (yes); outright switch, no dual-backend release; 1080p solid without macadam first, then UHD (and keying) as a separate plan.
- **Biggest risks:** native build/CI on both architectures (Task 2), and whether shared memory is enough to hit 24 fps at 1080p. If Task 6's measurement shows the remaining time is the `capturePage()` readback, the next step is offscreen rendering with shared textures, which is outside this plan.
