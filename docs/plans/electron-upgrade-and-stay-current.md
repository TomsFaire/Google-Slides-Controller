# Electron Upgrade & Stay-Current Design

**Date:** 2026-09-28
**Status:** Approved design — implementation plan not yet written
**Scope:** Two plans — (1) get from Electron 33.4.11 to 44.4.5, (2) a process that keeps us on a supported major.

---

## 1. Problem

The app ships **Electron 33.4.11** (Chromium 130.0.6723.191, Node 20.18.3), released 2025-04-26.
Current stable is **44.4.5** (Chromium 152.0.7977.130, Node 24.21.0), released 2026-09-23.

That is **11 major versions** and 22 Chromium majors behind.

Electron supports only the latest three majors. Electron 36 shipped 2025-04-28, which put the 33
line out of support two days after `33.4.11` was published — `33.4.11` was the final release of that
line. The app has therefore been shipping a **Chromium with no security patches for ~17 months**.

This matters more than the version number suggests: the app loads `docs.google.com` in a
`persist:google` session holding live Google credentials. An unpatched renderer is the actual risk.

Related drift found in the same audit:

- CI pins `node-version: '20'` ([.github/workflows/build.yml](../../.github/workflows/build.yml)). Node 20 reached end-of-life 2026-04-30.
- `electron-builder` is 24.13.3, published 2024-03-02. Current is 26.x.
- Root `yarn install --immutable` is a **no-op**: `--immutable` is a Yarn Berry flag and the root is `yarn@1.22.22`. Lockfile immutability has never been enforced at the repo root. (It *is* correct in the companion module, which is Yarn 4.)
- CI has **no test step**. The 103 unit tests in `tests/` have never run on a push or PR.
- No Dependabot or Renovate configuration exists.
- [INSTALLATION.md:74](../../INSTALLATION.md) claims "macOS 10.15+", already wrong for Electron 33, which requires macOS 11+.
- The endpoint table in [CLAUDE.md](../../CLAUDE.md) documents 10 routes and names two that do not exist (`/api/share-link`, `/api/show-share-qr`). `main.js` actually serves **40** routes; the real ones are `/api/show-tunnel-qr` and `/api/hide-tunnel-qr`.

---

## 2. Verified findings

### 2.1 The code is unusually well-positioned for this jump

Every breaking change in Electron's `docs/breaking-changes.md` from 34.0 through 44.0 was checked
against the app's actual API surface. The following are **no-ops for this codebase**:

| Breaking change | Why it does not apply |
|---|---|
| 44: `clipboard` module removed from renderer | `renderer.js` already uses `navigator.clipboard` (lines 935, 1190) |
| 44: `net.request` rejects frame destinations without navigate mode | The app's `net` is Node's TCP module ([main.js:25](../../main.js)), not Electron's `net` |
| 44: `webContents` may be `null` in `select-client-certificate` | Event not used |
| 44: workers in subframes need `nodeIntegrationInSubFrames` | `nodeIntegration: false` everywhere; no subframe Node use |
| 43: `NativeImage.toBitmap()` colour-space normalisation | `toBitmap()` not used; the app uses `capturePage().toDataURL()` |
| 42: offscreen rendering default device scale factor | No offscreen rendering |
| 41: PDFs no longer create a separate `WebContents` | No PDF handling |
| 38: `plugin-crashed` event removed | Event not used |
| 35: `WebRequestFilter.urls` empty-array semantics | The only filter ([main.js:1402](../../main.js)) passes an explicit non-empty array |
| 35: `setPreloads` / `getPreloads` deprecated | Not used |
| 35: `console-message` event argument deprecation | No listener |
| 32: `canGoBack` / `goBack` / `clearHistory` deprecations | Not used |
| 30: `BrowserView` deprecated | No `BrowserView` anywhere |
| 29: `crashed` / `renderer-process-crashed` removed | Already on `render-process-gone` |
| Native module ABI rebuilds | **Zero native modules** — `qrcode` and `png-to-ico` are pure JS; `node-pty` is an optional `try/catch` require that is not a declared dependency |

All 12 `webPreferences` blocks already use `nodeIntegration: false` and `contextIsolation: true`.

### 2.2 The three real risk sites

1. **GPU / ANGLE — [main.js:663-683](../../main.js).** Two changes compound: Electron 36 stopped
   propagating `app.commandLine` switches to child processes, and Electron 44 statically links ANGLE
   into the binary (loaded into every process; `libEGL`/`libGLESv2` no longer shipped). See §4 — this
   path is being removed rather than ported.
2. **`setWindowOpenHandler` — 8 call sites in `main.js`.** Electron 39 made `window.open` popups
   always resizable. Commit `e9f5524` recently reworked presenter-view window tracking by opener, so
   this area is already sensitive.
3. **macOS floor.** Electron 44 requires **macOS 13+**. Confirmed acceptable: all venue machines run
   macOS 13 or newer.

Electron's macOS floors across the range, for future reference: 33–37 require macOS 11+, 38–43
require macOS 12+, 44+ requires macOS 13+.

### 2.3 Toolchain constraints (verified against the npm registry)

- `electron@44.4.5` declares `engines.node: ">= 22.12.0"`. **CI's Node 20 is a hard build failure**, not a nice-to-have.
- The companion module ([companion-module-gslide-opener/package.json](../../companion-module-gslide-opener/package.json)) is a separate Yarn 4.12.0 workspace declaring `engines.node: "^22.20"`, i.e. `>=22.20.0 <23.0.0`. **Node 24 would violate it.**
- Therefore **Node 22.20+ is the only version satisfying both**, and it is LTS until 2027-04-30.
- `electron@44.4.5` and `electron-builder@26.x` both have **no `postinstall` script** (Electron 42 moved binary download to first run of the `bin` script; `ELECTRON_SKIP_BINARY_DOWNLOAD` is no longer supported).

---

## 3. Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Target version | **44.4.5** | All venue machines are macOS 13+, so nothing caps us below current stable |
| Node version | **22.20+** | Only version satisfying both `electron >= 22.12.0` and companion `^22.20` |
| Upgrade shape | **Toolchain first, then runtime** | Separates two independent failure domains for roughly the same total work; if the runtime PR goes wrong, revert one commit and keep a modern build chain |
| Verification | **Manual acceptance on real hardware, off-hours** | No Electron runtime tests are being built; the checklist in §6 is the deliverable |
| GPU mode path | **Remove it** (separate PR 3) | It was a diagnostic escape hatch added *during* the Electron 33 upgrade for a Chromium 130 transparent-PNG/compositing bug ([CHANGELOG.md:197](../../CHANGELOG.md)); no machine uses a non-default mode |
| Dependency automation | **PR for everything, never auto-merge** | User's explicit choice. §7 is designed to counter the queue-rot failure mode this creates |
| Version bump | **2.3.10 → 2.3.11** (patch) | Follows the patch-only rule. Open to 2.4.0 if preferred — see §8 |

---

## 4. Plan 1 — getting current

### PR 1 — Build modernisation (stays on Electron 33)

Goal: make the build chain current and provably working *before* the runtime moves, so PR 2 has
exactly one variable.

1. `.github/workflows/build.yml`: `node-version: '20'` → read from `.nvmrc` (`22.20`) via `node-version-file`.
2. Replace the no-op root `yarn install --immutable` with `yarn install --frozen-lockfile` (the Yarn 1 equivalent). Leave the companion module's `--immutable` alone — it is correct there.
3. `electron-builder` 24.13.3 → 26.x. Read the 24→25 and 25→26 changelogs and adapt the `build` block in [package.json](../../package.json) as needed; `afterPack`, `extraResources`, `electronLanguages` and `identity: null` are all areas electron-builder has changed.
4. Add `yarn test` to CI after install. The 103 tests currently never run; this becomes the only automated signal plan 2 has.
5. Confirm electron-builder 26 fetches Electron binaries correctly now that `postinstall` is gone.

**Verification (no venue rig):** CI green including the new test step; both mac zips build; extract
the arm64 `.app` and confirm it launches and opens the settings window.

**Explicitly excluded:** no Electron version change, no `main.js` changes. A failure here is
unambiguously the build chain.

### PR 2 — Electron 33.4.11 → 44.4.5

1. Bump `electron` to `^44.4.5` in [package.json](../../package.json); update `yarn.lock`.
2. Review the 8 `setWindowOpenHandler` sites against Electron 39's always-resizable popup change, with attention to presenter-view tracking (`e9f5524`).
3. Fix [INSTALLATION.md:74](../../INSTALLATION.md): macOS 10.15+ → macOS 13+.
4. Fix the stale endpoint table in [CLAUDE.md](../../CLAUDE.md) (40 real routes; `/api/share-link` and `/api/show-share-qr` do not exist).
5. `buildNumber` 88 → 89; version 2.3.10 → 2.3.11; CHANGELOG entry.

**Verification:** the full §6 checklist on real hardware. GPU mode is verified in **`default` only** —
the other four modes are used by no machine and PR 3 deletes them, so testing them would spend a
hardware session on dead code.

### PR 3 — Remove the GPU mode path

Licensed by checklist item 1.3 passing (transparent PNG / layered slides render correctly on
Chromium 152 in `default` mode).

Remove:
- [index.html:652-660](../../index.html) — the `<select>` and its hint text
- The 6 `presentationGpuMode` references in [renderer.js](../../renderer.js) (element binding, load, save)
- [main.js:524](../../main.js) — `VALID_PRESENTATION_GPU_MODES`
- [main.js:663-683](../../main.js) — `applyPresentationGpuCommandLineEarly()` and its call
- The `presentationGpuMode` normalisation in `loadPreferences()`
- The `GSLIDE_GPU_MODE` branch in [scripts/slides-presenter-gpu-smoke.cjs](../../scripts/slides-presenter-gpu-smoke.cjs)

**Keep:** `presentationGpuMode` in the `/api/status` response, as a frozen `'default'` literal.
[CHANGELOG.md:200](../../CHANGELOG.md) documents it as part of the API contract and the HTTP API
exists for AV integrations we do not control. Keeping the field costs nothing and avoids a breaking
response change.

**Keep:** `presentationNativeFullscreen`. Unlike GPU mode it is a genuine behavioural choice
(`setSimpleFullScreen` vs `setFullScreen`) and some AV rigs need native fullscreen. Removing it is a
separate conversation.

**Verification:** light pass — app launches, presents, slides render correctly.

---

## 5. Out of scope

- Auto-merging any dependency update (user's explicit choice).
- A Renovate app installation (Dependabot plus §7.3 covers the need).
- An Electron runtime smoke-test harness in CI (manual hardware verification was chosen instead).
- Expanding CI to build Windows or Linux targets.
- Removing `presentationNativeFullscreen`.
- Any unrelated refactoring of `main.js`, despite its size.

---

## 6. Hardware acceptance checklist

Design principle: only test what an Electron/Chromium jump can break. The 103 unit tests already
cover the pure-JS logic (PerfectCue parsing, key-fill, preset serialisation). Ordered to fail fast.

### Gate 0 — dev machine, before touching the rig
- [ ] `yarn test` green
- [ ] App launches; settings window renders
- [ ] `curl -s localhost:9595/api/status | jq .runtime` reports **chrome 152 / electron 44 / node 24.21** — proves the bump took, not a cached binary.
      Note: this `node` value is the Node **bundled inside Electron 44** (24.21.0), which is a different thing from the Node that *builds* the app (22.20, §2.3). Both numbers are correct.

### Gate 1 — Chromium 130 to 152 (highest risk; do first)
- [ ] Google login in `persist:google`, then **quit and relaunch** — session persists
- [ ] Open a deck, send to second monitor, fullscreen — correct monitor, no letterboxing
- [ ] **Transparent PNG / layered slides, side by side against Chrome** — the bug the GPU selector existed for. Passing this licenses PR 3
- [ ] Speaker notes render with **no U+FFFD (replacement character) corruption** — exercises `normalizeSpeakerNotes()` and the `onHeadersReceived` charset rewrite at [main.js:1402](../../main.js). See [docs/SPEAKER-NOTES-ENCODING.md](../SPEAKER-NOTES-ENCODING.md)
- [ ] Presenter view tracked by opener — confirms `e9f5524` survives Electron 39's popup change
- [ ] `/api/get-slide-previews` thumbnails non-blank and correctly scaled (`capturePage()`)
- [ ] Notes scroll and zoom endpoints — these `executeJavaScript` against Google's live DOM

### Gate 2 — Electron API behaviour changes
- [ ] Preset export/import dialogs — Electron 43 changed the default directory to Downloads; confirm a sane location and a clean round-trip
- [ ] Key-fill window opens and closes (own session partition)
- [ ] Slido window opens (own session partition)
- [ ] Stagetimer overlay show / hide / update settings — the largest single feature (430 references in `main.js`)
- [ ] Cloudflared tunnel enable, then `/api/show-tunnel-qr` overlay renders on the presentation screen
- [ ] `safeStorage` round-trip for the Cloudflare token (macOS keychain; `isEncryptionAvailable()`)

### Gate 3 — smoke only (not Chromium-sensitive)
- [ ] `/api/status`, `/api/displays`, `/api/presets` respond; IP allowlist rejects a disallowed source
- [ ] PerfectCue TCP connect plus one next/prev (Node `net`, unaffected — sanity only)
- [ ] `sendToBackups()` reaches a backup; `/api/backup-status` reports correctly

---

## 7. Plan 2 — staying current

The constraint: Electron ships a major roughly every 8 weeks and supports the latest three, so each
major grants about **24 weeks of runway**. The process must force a move at least every ~5 months.

Because every update requires manual review and nothing auto-merges, the failure mode to design
against is not a bad merge — it is a **queue that silently rots**, which is how the app reached 17
months behind. Parts 3 and 4 exist specifically to counter that.

### 7.1 Single source of truth for the Node floor
Add `engines.node: ">=22.20 <23"` to root [package.json](../../package.json) and a `.nvmrc`
containing `22.20`. CI reads `.nvmrc` via `node-version-file`. The floor currently lives in three
disagreeing places, which is how Node 20 survived until it became a hard blocker.

### 7.2 Dependabot, weekly, grouped
A `.github/dependabot.yml` covering both ecosystems — root (Yarn 1) and
`companion-module-gslide-opener` (Yarn 4) — with `electron` and `electron-builder` in a single group
so the coupled pair always moves together. Every bump opens a PR and waits for review.

### 7.3 Electron support-drift alarm (the forcing function)
A weekly scheduled workflow, `.github/workflows/electron-support-check.yml`, that compares the
`electron` version in `package.json` against Electron's published releases and escalates:

| State | Action |
|---|---|
| Our major is within the latest 3 | Exit silently — zero noise |
| Fell to N-3 (just lost support) | Open **one** issue labelled `electron-drift`, titled with the delta |
| 2+ majors past end-of-life | Update that same issue, escalate the label, **and fail the workflow** (visible red X on the Actions tab) |

The anti-rot property is that it maintains **exactly one** issue, located by label and updated in
place, so it can never decay into a pile of stale notifications. Its body carries the actionable
delta: current vs. target major, the target's **macOS floor**, and the target's **`engines.node`**
requirement — the facts that had to be gathered by hand for this document.

Data source: `https://releases.electronjs.org/releases.json` or the npm `dist-tags` endpoint.

### 7.4 The checklist becomes a repo artifact
§6 lands at `docs/electron-upgrade-checklist.md`, together with the **method** used here:

> Diff Electron's `docs/breaking-changes.md` between the current and target versions, then grep the
> app for every API each entry names. Record the no-ops as well as the hits — the no-op table is
> what shrinks an 11-major jump to a three-item risk list.

This is the highest-leverage item in plan 2: it converts each future upgrade from archaeology into a
checklist run.

### 7.5 Cadence policy
**Upgrade while still supported, not after end-of-life.** Move on reaching N-2, which is roughly
twice a year and 2-3 Electron majors per hop. The policy document names a **blackout window**
(filled in with the event season) so a Chromium jump never lands the week of a show.

---

## 8. Open parameters

1. **Version bump size.** Defaulting to patch (2.3.10 → 2.3.11) per the patch-only rule. A jump of
   22 Chromium majors arguably justifies a minor (2.4.0). Owner's call.
2. **Blackout window.** §7.5 needs the event season filled in before the policy is written down.
