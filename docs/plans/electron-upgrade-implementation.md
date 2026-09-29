# Electron 33 → 44 Upgrade Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move the app from Electron 33.4.11 (unsupported since 2025-04-28) to 44.4.5, then delete the GPU workaround that Chromium 152 makes obsolete.

**Architecture:** Three sequential pull requests, each with its own verification gate. PR 1 modernises the build toolchain while staying on Electron 33, so a failure there is unambiguously the build chain. PR 2 changes only the runtime, on a toolchain already proven good. PR 3 removes the GPU mode path once the hardware pass has proven it obsolete.

**Tech Stack:** Electron, electron-builder, Yarn 1 (root) / Yarn 4 (companion module), GitHub Actions, `node:test`.

**Spec:** [docs/plans/electron-upgrade-and-stay-current.md](electron-upgrade-and-stay-current.md) — read it alongside this plan; §2.1 lists the breaking changes that are deliberately *not* addressed because they are no-ops for this codebase.

## Global Constraints

- **Target Electron:** `^44.4.5` exactly. Not 45.x — 45 is alpha at time of writing.
- **Node floor:** `>=22.20 <23`. This is the *only* range satisfying both `electron@44.4.5` (`engines.node: ">= 22.12.0"`) and the companion module (`engines.node: "^22.20"`). **Node 24 would violate the companion module.**
- **electron-builder:** `^26.15.3` — the version the npm `latest` tag points to.
- **macOS floor:** Electron 44 requires macOS 13+. All venue machines are confirmed macOS 13+.
- **Version bump:** 2.3.10 → **2.3.11** (patch), `buildNumber` 88 → **89**. Per the repo's patch-only rule.
- **Keep** `presentationNativeFullscreen` everywhere. Only `presentationGpuMode` is being removed.
- **Keep** `presentationGpuMode` in the `/api/status` response as a frozen `'default'` literal — it is a documented API contract ([CHANGELOG.md:200](../../CHANGELOG.md)).
- **No unrelated refactoring** of `main.js`, despite its size.

**A note on testing in this plan.** Most tasks here change configuration, versions, or delete code — there is no new logic to drive with a unit test, and writing one would be theatre. Those tasks end with a concrete build-or-run verification command and its expected output instead. `tests/` gains no new files in this plan; the genuine TDD work lives in the companion plan ([electron-stay-current-implementation.md](electron-stay-current-implementation.md)), where the drift alarm has real logic. Note also that `main.js` cannot be `require`d from a test — it calls `require('electron')` at module scope — which is why the repo's own tests either import from `src/` or inline a copy (see the comment at the top of [tests/key-fill.test.js](../../tests/key-fill.test.js)).

---

## File Structure

| File | Change | Responsibility after change |
|---|---|---|
| `.nvmrc` | **create** | Single source of truth for the Node version; read by CI and `nvm use` |
| `package.json` | modify | Adds `engines.node`; bumps `electron` and `electron-builder`; version + buildNumber |
| `.github/workflows/build.yml` | modify | Reads Node from `.nvmrc`; enforces the lockfile correctly; runs the test suite |
| `main.js` | modify | Loses the GPU switch application, its constant, and two prefs-normalisation blocks (PR 3) |
| `renderer.js` | modify | Loses the GPU select binding and its save path (PR 3) |
| `index.html` | modify | Loses the GPU mode form group (PR 3) |
| `scripts/slides-presenter-gpu-smoke.cjs` | modify | Loses the `GSLIDE_GPU_MODE` branch; keeps the session-partition option (PR 3) |
| `INSTALLATION.md` | modify | Correct macOS floor |
| `CLAUDE.md` | modify | Correct endpoint table |
| `CHANGELOG.md` | modify | 2.3.11 entry |
| `docs/electron-upgrade-checklist.md` | **create** | Reusable acceptance checklist + the breaking-change audit method |

---

# PR 1 — Build modernisation (stays on Electron 33)

Branch: `chore/toolchain-node22-builder26`

## Task 1: Single source of truth for the Node version

**Why first:** `electron@44.4.5` will not install on Node 20, so this unblocks PR 2. It also fixes the root cause of the drift — the Node floor currently lives in three disagreeing places.

**Files:**
- Create: `.nvmrc`
- Modify: `package.json` (add `engines`)
- Modify: `.github/workflows/build.yml:55-59`

- [ ] **Step 1: Create `.nvmrc`**

```
22.20.0
```

An exact patch version, so CI and local installs are deterministic. Raising it is a deliberate commit, not a surprise.

- [ ] **Step 2: Add `engines` to root `package.json`**

Insert after the `"license": "MIT",` line:

```json
  "engines": {
    "node": ">=22.20 <23"
  },
```

The upper bound is not cosmetic: the companion module declares `^22.20`, so Node 23+ would break it.

- [ ] **Step 3: Point CI at `.nvmrc`**

In `.github/workflows/build.yml`, replace lines 55-59:

```yaml
      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '20'
          cache: 'yarn'
```

with:

```yaml
      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version-file: '.nvmrc'
          cache: 'yarn'
```

- [ ] **Step 4: Verify locally**

```bash
export NVM_DIR="$HOME/.nvm" && source "$NVM_DIR/nvm.sh"
nvm install && nvm use
node --version
```

Expected: `v22.20.0`

- [ ] **Step 5: Verify the engine range is satisfiable**

```bash
node -e "const s=require('semver');" 2>/dev/null || true
node -p "process.version"
```

Expected: a `v22.20.x` string. If you are on Node 20 here, stop — later tasks will fail confusingly.

- [ ] **Step 6: Commit**

```bash
git add .nvmrc package.json .github/workflows/build.yml
git commit -m "chore: pin Node to 22.20 via .nvmrc and engines

electron@44.4.5 requires node >=22.12.0 and the companion module
requires ^22.20, so 22.20 is the only range satisfying both. CI was
pinned to Node 20, which reached end-of-life 2026-04-30."
```

---

## Task 2: Make CI enforce the lockfile and run the tests

**Why:** the root `yarn install --immutable` is a **no-op** — `--immutable` is a Yarn Berry flag and the root is `yarn@1.22.22`. Separately, the 103 tests in `tests/` have never run in CI. This task is the only automated signal the stay-current plan will have.

**Files:**
- Modify: `.github/workflows/build.yml:61-62` (install), plus a new step after it

**Interfaces:**
- Produces: a CI job that fails on a stale lockfile or a failing test. The stay-current plan's Dependabot PRs depend on this being real.

- [ ] **Step 1: Confirm the tests currently pass**

```bash
yarn test
```

Expected: all tests pass. Note the count — there are 103 `test(` calls across 7 files. If anything fails *before* you change a thing, stop and report it; do not fold a pre-existing failure into this PR.

- [ ] **Step 2: Fix the root install to actually enforce the lockfile**

In `.github/workflows/build.yml`, replace line 62:

```yaml
        run: yarn install --immutable
```

with:

```yaml
        run: yarn install --frozen-lockfile
```

**Leave line 124 alone.** That step runs in `./companion-module-gslide-opener`, which is Yarn 4, where `--immutable` is the correct flag.

- [ ] **Step 3: Add the test step**

Immediately after the `Install dependencies` step, insert:

```yaml
      - name: Run tests
        run: yarn test
```

- [ ] **Step 4: Verify the no-op claim, so the change is justified in review**

```bash
yarn install --immutable 2>&1 | tail -5
```

Expected: Yarn 1 completes normally and does **not** error on the unknown flag — demonstrating it was never enforcing anything.

- [ ] **Step 5: Verify the replacement flag works**

```bash
yarn install --frozen-lockfile
```

Expected: completes successfully with the lockfile unchanged (`git status` shows no change to `yarn.lock`).

- [ ] **Step 6: Commit**

```bash
git add .github/workflows/build.yml
git commit -m "ci: enforce the lockfile and run the test suite

Root is yarn@1.22.22, where --immutable is not a recognised flag, so
lockfile enforcement has been silently absent. --frozen-lockfile is the
Yarn 1 equivalent. The companion module's --immutable is correct and is
left as-is. Also runs the 103 existing tests, which CI never ran."
```

---

## Task 3: electron-builder 24.13.3 → 26.x

**Why:** electron-builder 24.13.3 predates Electron 44 by over two years. This is the task most likely to need judgement, so it is isolated in its own commit.

**Files:**
- Modify: `package.json` (`devDependencies.electron-builder`)
- Possibly modify: `package.json` (the `build` block)

- [ ] **Step 1: Read the changelogs before changing anything**

The `build` block in `package.json` uses `afterPack`, `extraResources`, `electronLanguages`, `identity: null`, `signingHashAlgorithms: []`, and `publish: null` — all areas electron-builder has touched between 24 and 26. Check each against the 25.0.0 and 26.0.0 notes:

```bash
open https://github.com/electron-userland/electron-builder/releases
```

Note: the npm `latest` tag resolves to `26.15.3` while a `v26` tag points at `26.17.0`. Use **`^26.15.3`** — follow the `latest` tag rather than the higher number.

- [ ] **Step 2: Bump the dependency**

In `package.json`, change:

```json
    "electron-builder": "^24.13.3",
```

to:

```json
    "electron-builder": "^26.15.3",
```

- [ ] **Step 3: Install**

```bash
yarn install
```

Expected: `yarn.lock` updates. Confirm the resolved version:

```bash
node -p "require('electron-builder/package.json').version"
```

Expected: a `26.x.y` string.

- [ ] **Step 4: Build, still on Electron 33**

```bash
yarn build:mac
```

Expected: `dist/mac-arm64/Google Slides Opener.app` exists. If the build fails, the cause is in the `build` block — fix it here, and record what changed in the commit message.

- [ ] **Step 5: Launch the packaged app**

```bash
open "dist/mac-arm64/Google Slides Opener.app"
```

Expected: the settings window appears. (If the sandbox blocks `open`, run it yourself from Finder — the launch is the verification and must not be skipped.)

- [ ] **Step 6: Confirm the runtime is still Electron 33**

With the app running:

```bash
curl -s localhost:9595/api/status | python3 -m json.tool | grep -A4 runtime
```

Expected: `electron` reports `33.4.11`. This task must not move the runtime.

- [ ] **Step 7: Commit**

```bash
git add package.json yarn.lock
git commit -m "build: upgrade electron-builder 24.13.3 -> 26.x

Still on Electron 33, so any packaging regression here is attributable
to the builder rather than the runtime."
```

---

## Gate A — PR 1 verification

- [ ] CI is green on the PR, **including the new test step**
- [ ] Both macOS zips are produced as artifacts (arm64 and x64)
- [ ] The extracted arm64 `.app` launches and opens the settings window
- [ ] `/api/status` still reports Electron `33.4.11`
- [ ] Open the PR with the `claude` label; merge only once green

---

# PR 2 — Electron 33.4.11 → 44.4.5

Branch: `chore/electron-44`

## Task 4: Bump Electron

**Files:**
- Modify: `package.json` (`devDependencies.electron`)
- Modify: `yarn.lock`

- [ ] **Step 1: Bump the dependency**

In `package.json`, change:

```json
    "electron": "^33.4.11",
```

to:

```json
    "electron": "^44.4.5",
```

- [ ] **Step 2: Install and confirm the resolved version**

```bash
yarn install
node -p "require('electron/package.json').version"
```

Expected: `44.4.5`

Note: Electron 42 removed the `postinstall` binary download. The binary is now fetched on first run of the `bin` script. If `node_modules/electron/dist` is absent, that is expected — it populates on first launch, or on demand via `npx install-electron --no`.

- [ ] **Step 3: Fast boot check before involving the rig**

```bash
yarn smoke:slides-gpu "https://docs.google.com/presentation/d/<any-deck-id>/present"
```

Expected: a window opens and renders the deck. This is a 30-second sanity check that Electron 44 boots at all — do it before booking hardware time.

- [ ] **Step 4: Confirm the bundled runtime versions**

```bash
yarn start
```

then, with the app running:

```bash
curl -s localhost:9595/api/status | python3 -m json.tool | grep -A4 runtime
```

Expected: `chrome` reports `152.x`, `electron` reports `44.4.5`, `node` reports `24.21.x`.

The `node` value here is the Node **bundled inside Electron 44**, which is deliberately different from the Node that *builds* the app (22.20, Task 1). Both numbers are correct; do not "fix" either.

- [ ] **Step 5: Run the test suite**

```bash
yarn test
```

Expected: all pass. These are pure-JS tests, so a failure here means something unexpected — investigate before continuing.

- [ ] **Step 6: Commit**

```bash
git add package.json yarn.lock
git commit -m "feat: upgrade Electron 33.4.11 -> 44.4.5

Chromium 130 -> 152, 11 major versions. The 33 line lost support on
2025-04-28 when Electron 36 shipped; 33.4.11 was its final release."
```

---

## Task 5: Review `setWindowOpenHandler` against Electron 39

**Why:** Electron 39 made `window.open` popups **always resizable**. Commit `e9f5524` recently reworked presenter-view window tracking by opener, so popup handling is already delicate here.

**Files:**
- Review, and modify only if needed: `main.js` at lines 2448, 2548, 2686, 3617, 3701, 4524, 4656, 6214

- [ ] **Step 1: Inspect all eight call sites**

```bash
grep -n "setWindowOpenHandler" -A12 main.js
```

- [ ] **Step 2: Identify any handler that depends on a popup being non-resizable**

For each site, check whether the returned object sets `overrideBrowserWindowOptions` with `resizable: false`, or whether downstream code assumes fixed popup geometry. Electron 39 ignores a `resizable: false` request for `window.open` popups, so any such assumption is now wrong.

- [ ] **Step 3: If a site relied on fixed geometry, pin the size explicitly instead**

Replace reliance on non-resizability with explicit bounds after creation, e.g.:

```javascript
presentationWindow.webContents.setWindowOpenHandler(({ url }) => ({
  action: 'allow',
  overrideBrowserWindowOptions: { width: 1280, height: 720 }
}));
```

If no site relies on it — the likely outcome, since the handlers mostly return a bare `{ action: 'allow' }` — make **no code change** and record that finding in the commit for Task 8. Do not invent a change to look busy.

- [ ] **Step 4: Verify presenter view behaviour**

```bash
yarn start
```

Open a deck with speaker notes, trigger presenter view, and confirm the notes window is adopted (this is what `src/notes-window-adoption.js` handles).

- [ ] **Step 5: Run the adoption tests**

```bash
node --test tests/notes-window-adoption.test.js
```

Expected: PASS.

- [ ] **Step 6: Commit only if you changed something**

```bash
git add main.js
git commit -m "fix: pin popup geometry explicitly for Electron 39

Electron 39 makes window.open popups always resizable, so requesting
resizable: false no longer has any effect."
```

---

## Task 6: Correct the stale documentation

**Files:**
- Modify: `INSTALLATION.md:74`
- Modify: `CLAUDE.md` (the API endpoint table)

- [ ] **Step 1: Fix the macOS floor**

In `INSTALLATION.md`, change line 74 from:

```
- macOS 10.15+
```

to:

```
- macOS 13+ (Ventura or newer — required by Electron 44)
```

The old value was already wrong: Electron 33 dropped macOS 10.15 support.

- [ ] **Step 2: Get the real route list**

```bash
grep -oE "'/api/[a-z0-9-]+'" main.js | tr -d "'" | sort -u
```

Expected: 40 routes.

- [ ] **Step 3: Correct the endpoint table in `CLAUDE.md`**

Two documented routes do not exist and must be removed or replaced: `/api/share-link` and `/api/show-share-qr`. The real tunnel-QR routes are `/api/show-tunnel-qr` and `/api/hide-tunnel-qr`. Replace those two rows:

```markdown
| POST | `/api/show-tunnel-qr` | Display tunnel QR code on presentation screen |
| POST | `/api/hide-tunnel-qr` | Hide the tunnel QR overlay |
```

Then add a line under the table so it stops drifting:

```markdown
The table above covers the commonly used routes. `main.js` serves 40 in total; run
`grep -oE "'/api/[a-z0-9-]+'" main.js | tr -d "'" | sort -u` for the authoritative list.
```

- [ ] **Step 4: Verify no other doc claims a pre-13 macOS floor**

```bash
grep -rn "10\.15\|macOS 11\|macOS 12\|Monterey\|Big Sur" README.md INSTALLATION.md RELEASES.md docs/*.md
```

Expected: no remaining claims of a macOS floor below 13. Fix any that appear.

- [ ] **Step 5: Commit**

```bash
git add INSTALLATION.md CLAUDE.md
git commit -m "docs: correct macOS floor and API endpoint table

macOS 10.15+ was already wrong for Electron 33 and is now 13+ under
Electron 44. The endpoint table named two routes that do not exist
(/api/share-link, /api/show-share-qr) and documented 10 of 40 routes."
```

---

## Task 7: Commit the acceptance checklist as a repo artifact

**Why:** this converts every future Electron upgrade from archaeology into a checklist run. It must exist *before* the hardware pass, so the pass is run from the committed artifact rather than from chat scrollback.

**Files:**
- Create: `docs/electron-upgrade-checklist.md`

- [ ] **Step 1: Create the file**

```markdown
# Electron Upgrade Acceptance Checklist

Run this on real hardware for every Electron major upgrade. Ordered to fail fast:
if Gate 1 fails, stop and fix before spending time on Gates 2 and 3.

Design principle: only test what an Electron/Chromium jump can break. The unit
tests in `tests/` already cover the pure-JS logic (PerfectCue parsing, key-fill,
preset serialisation); re-testing those by hand wastes the scarce resource.

## Gate 0 — dev machine, before touching the rig

- [ ] `yarn test` green
- [ ] App launches; settings window renders
- [ ] `curl -s localhost:9595/api/status | python3 -m json.tool | grep -A4 runtime`
      reports the expected Chromium / Electron / Node versions — proves the bump
      took effect rather than a cached binary.
      Note: the `node` value is the Node bundled *inside* Electron, which is not
      the Node that builds the app (see `.nvmrc`). Both differing is correct.

## Gate 1 — Chromium jump (highest risk; do first)

- [ ] Google login in `persist:google`, then quit and relaunch — session persists
- [ ] Open a deck, send to second monitor, fullscreen — correct monitor, no letterboxing
- [ ] Transparent PNG / layered slides render correctly, compared side by side against Chrome
- [ ] Speaker notes render with no U+FFFD (replacement character) corruption —
      exercises `normalizeSpeakerNotes()` and the `onHeadersReceived` charset
      rewrite. See docs/SPEAKER-NOTES-ENCODING.md
- [ ] Presenter view window is adopted by opener (see `src/notes-window-adoption.js`)
- [ ] `/api/get-slide-previews` thumbnails are non-blank and correctly scaled (`capturePage()`)
- [ ] Notes scroll and zoom endpoints work — these `executeJavaScript` against Google's live DOM

## Gate 2 — Electron API behaviour changes

- [ ] Preset export/import dialogs open somewhere sane and round-trip cleanly
      (Electron 43 changed the default directory to Downloads)
- [ ] Key-fill window opens and closes (own session partition)
- [ ] Slido window opens (own session partition)
- [ ] Stagetimer overlay show / hide / update settings — the largest single feature
- [ ] Cloudflared tunnel enable, then `/api/show-tunnel-qr` renders on the presentation screen
- [ ] `safeStorage` round-trip for the Cloudflare token (macOS keychain)

## Gate 3 — smoke only (not Chromium-sensitive)

- [ ] `/api/status`, `/api/displays`, `/api/presets` respond; IP allowlist rejects a disallowed source
- [ ] PerfectCue TCP connect plus one next/prev (Node `net`, unaffected — sanity only)
- [ ] `sendToBackups()` reaches a backup; `/api/backup-status` reports correctly

---

## Method: auditing breaking changes before an upgrade

This is what turns "N majors of unknown risk" into a short, verified risk list.

1. Fetch the target version's breaking-change document:

       curl -s -o /tmp/bc.md \
         https://raw.githubusercontent.com/electron/electron/v<TARGET>/docs/breaking-changes.md

2. List the section headers between your version and the target:

       grep -nE "^## |^### " /tmp/bc.md

3. For each entry, grep the app for the API it names — `main.js`, `preload.js`,
   `renderer.js`. Record the misses as well as the hits.
4. The table of no-ops is the valuable output: it is what lets review focus on the
   handful of real risk sites instead of re-reading the whole diff.

Two traps found doing this for 33 → 44, worth checking every time:

- Node's `net` and Electron's `net` are different modules. Check which one
  `require('net')` refers to before assuming an Electron `net` change applies.
- Check each version's **macOS floor** and **`engines.node`** early; both can cap
  the target independently of any code concern.
```

- [ ] **Step 2: Verify the internal links resolve**

```bash
cd docs && for f in SPEAKER-NOTES-ENCODING.md; do [ -e "$f" ] && echo "OK $f" || echo "MISS $f"; done; cd ..
```

Expected: `OK SPEAKER-NOTES-ENCODING.md`

- [ ] **Step 3: Commit**

```bash
git add docs/electron-upgrade-checklist.md
git commit -m "docs: add reusable Electron upgrade acceptance checklist

Records both the checklist and the breaking-change audit method, so the
next upgrade is a checklist run rather than an archaeology exercise."
```

---

## Task 8: Version, build number, changelog

**Files:**
- Modify: `package.json` (`version`, `buildNumber`)
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Bump version and build number**

In `package.json`: `"version": "2.3.10"` → `"2.3.11"`, and `"buildNumber": "88"` → `"89"`.

The build number must be incremented before any packaging build.

- [ ] **Step 2: Add the changelog entry**

At the top of `CHANGELOG.md`, following the file's existing format:

```markdown
## 2.3.11

- **Electron 33.4.11 → 44.4.5** (Chromium 130 → 152, bundled Node 20 → 24). The
  Electron 33 line lost support on 2025-04-28, so the bundled Chromium had been
  unpatched for ~17 months. **Requires macOS 13+.**
- **Build toolchain:** electron-builder 24 → 26; CI Node 20 → 22.20 (Node 20
  reached end-of-life 2026-04-30); CI now enforces the lockfile correctly and runs
  the test suite.
- **Docs:** corrected the macOS floor in INSTALLATION.md and the API endpoint table
  in CLAUDE.md; added docs/electron-upgrade-checklist.md.
```

If Task 5 required a code change, add a line for it. If it did not, say so explicitly — a reviewer should not have to wonder whether the popup change was considered.

- [ ] **Step 3: Verify**

```bash
node -p "const p=require('./package.json'); p.version + ' build ' + p.buildNumber"
```

Expected: `2.3.11 build 89`

- [ ] **Step 4: Commit**

```bash
git add package.json CHANGELOG.md
git commit -m "chore: release 2.3.11 (build 89)"
```

---

## Gate B — PR 2 verification

- [ ] CI green, including the test step
- [ ] `yarn build:mac` produces both architectures
- [ ] **Run the full `docs/electron-upgrade-checklist.md` on real hardware, off-hours**
- [ ] Gate 1 item 3 (transparent PNG / layered slides on `default` GPU mode) **passed** — this is what licenses PR 3. If it failed, stop: PR 3 must not proceed, and the GPU path stays.
- [ ] Open the PR with the `claude` label; merge only after the hardware pass

---

# PR 3 — Remove the GPU mode path

Branch: `chore/remove-gpu-mode-path`

**Precondition:** Gate B's transparent-PNG check passed on `default` mode. Do not start otherwise.

## Task 9: Delete the GPU mode selector and its plumbing

**Why:** the selector was a diagnostic escape hatch added *during* the Electron 33 upgrade for a Chromium 130 transparent-PNG/compositing bug ([CHANGELOG.md:197](../../CHANGELOG.md)). No machine uses a non-default mode, and Chromium 152 has been verified to render correctly without it.

**Files:**
- Modify: `index.html:651-661` (the form group)
- Modify: `renderer.js:555-560` and `renderer.js:1596-1608`
- Modify: `main.js:524`, `main.js:658-685`, `main.js:1462-1463`, `main.js:3137-3139`
- Modify: `scripts/slides-presenter-gpu-smoke.cjs`
- **Do not touch:** `main.js:3944`

- [ ] **Step 1: Remove the UI form group**

In `index.html`, delete this entire block (lines 651-661), leaving the surrounding `Presentation Rendering` panel and its native-fullscreen control intact:

```html
            <div class="form-group">
              <label for="presentation-gpu-mode">GPU mode</label>
              <select id="presentation-gpu-mode" class="select-input">
                <option value="default">Default (Chromium decides)</option>
                <option value="angle-metal">ANGLE Metal (macOS)</option>
                <option value="angle-gl">ANGLE OpenGL</option>
                <option value="swiftshader">SwiftShader (software GL)</option>
                <option value="disable-gpu">Disable GPU (software compositing)</option>
              </select>
              <p class="field-hint">If slide images show wrong transparency, boxes, or seams vs Chrome, try ANGLE Metal or Disable GPU. <strong>Restart the app</strong> after changing this.</p>
            </div>
```

- [ ] **Step 2: Remove the renderer load binding**

In `renderer.js`, delete lines 555-560:

```javascript
    const presentationGpuModeSelect = document.getElementById('presentation-gpu-mode');
    if (presentationGpuModeSelect) {
      const allowedGpu = ['default', 'angle-metal', 'angle-gl', 'swiftshader', 'disable-gpu'];
      const m = String(preferences.presentationGpuMode || 'default');
      presentationGpuModeSelect.value = allowedGpu.includes(m) ? m : 'default';
    }
```

- [ ] **Step 3: Rewrite the renderer save path**

In `renderer.js`, replace the body of `savePresentationRenderingPreferences()` (lines 1596-1608). It currently saves both settings and its toast mentions GPU mode, which will no longer be true:

```javascript
async function savePresentationRenderingPreferences() {
  try {
    const nativeEl = document.getElementById('presentation-native-fullscreen');
    await window.electronAPI.savePreferences({
      presentationNativeFullscreen: nativeEl ? nativeEl.checked === true : false
    });
    showStatus('Presentation rendering saved.', 'info');
  } catch (error) {
```

Leave the existing `catch` block below it unchanged.

- [ ] **Step 4: Remove the main-process constant**

In `main.js`, delete line 524:

```javascript
const VALID_PRESENTATION_GPU_MODES = new Set(['default', 'disable-gpu', 'angle-metal', 'angle-gl', 'swiftshader']);
```

- [ ] **Step 5: Remove the switch application and its call**

In `main.js`, delete the `applyPresentationGpuCommandLineEarly()` function (lines ~658-683, including its leading doc comment) and the bare `applyPresentationGpuCommandLineEarly();` call immediately after it.

Leave `wantsPresentationNativeFullscreen()` and `applyPresentationFullscreenChrome()`, which follow it, completely intact.

- [ ] **Step 6: Remove the two prefs-normalisation blocks**

In `main.js`, delete lines 1462-1463:

```javascript
      if (!VALID_PRESENTATION_GPU_MODES.has(String(prefs.presentationGpuMode || ''))) {
        prefs.presentationGpuMode = 'default';
      }
```

and lines 3137-3139:

```javascript
  if (incoming && incoming.presentationGpuMode !== undefined) {
    const g = String(incoming.presentationGpuMode || '');
    mergedPrefs.presentationGpuMode = VALID_PRESENTATION_GPU_MODES.has(g) ? g : 'default';
  }
```

Mind the surrounding braces — both sit inside larger `if`/function bodies.

- [ ] **Step 7: Freeze the API field**

In `main.js` line 3944, change:

```javascript
          presentationGpuMode: statusPrefs.presentationGpuMode || 'default',
```

to:

```javascript
          // Retained as a frozen literal: documented API contract (CHANGELOG 1.9.12),
          // consumed by AV integrations we do not control. The GPU mode selector
          // itself was removed in 2.3.12 once Chromium 152 made it unnecessary.
          presentationGpuMode: 'default',
```

- [ ] **Step 8: Simplify the smoke script**

In `scripts/slides-presenter-gpu-smoke.cjs`, delete `applyGpuModeFromEnv()` and its call, and remove the `GSLIDE_GPU_MODE` line from the usage comment. Keep `GSLIDE_SESSION_PARTITION`, which is still useful.

- [ ] **Step 9: Verify only the frozen literal remains**

```bash
grep -rn "presentationGpuMode\|VALID_PRESENTATION_GPU_MODES\|GSLIDE_GPU_MODE\|presentation-gpu-mode" main.js renderer.js index.html preload.js scripts/
```

Expected: exactly **one** hit — the frozen literal and its comment at `main.js:~3944`. Any other hit is an incomplete removal.

- [ ] **Step 10: Verify nothing else referenced it**

```bash
grep -rn "presentationGpuMode" --include="*.js" --include="*.html" . 2>/dev/null | grep -v node_modules | grep -v "^./docs/"
```

Expected: the same single hit. `docs/` matches are historical records and should be left alone.

- [ ] **Step 11: Run the tests**

```bash
yarn test
```

Expected: all pass.

- [ ] **Step 12: Verify the app and the API contract**

```bash
yarn start
```

Then:

```bash
curl -s localhost:9595/api/status | python3 -c "import json,sys; print(json.load(sys.stdin)['presentationGpuMode'])"
```

Expected: `default`

Also open Settings → the `Presentation Rendering` panel: the GPU dropdown is gone, the native-fullscreen checkbox still works, and saving it shows "Presentation rendering saved."

- [ ] **Step 13: Bump version and changelog**

`package.json`: version → `2.3.12`, `buildNumber` → `90`. In `CHANGELOG.md`:

```markdown
## 2.3.12

- **Removed the Presentation GPU mode setting.** It was a diagnostic workaround for a
  Chromium 130 transparent-PNG/compositing bug, added during the Electron 33 upgrade.
  Chromium 152 renders these decks correctly, verified on hardware. `/api/status` still
  reports `presentationGpuMode: "default"` so existing AV integrations are unaffected.
  macOS native fullscreen is unchanged.
```

- [ ] **Step 14: Commit**

```bash
git add index.html renderer.js main.js scripts/slides-presenter-gpu-smoke.cjs package.json CHANGELOG.md
git commit -m "refactor: remove the Presentation GPU mode workaround

Added during the Electron 33 upgrade to work around a Chromium 130
transparent-PNG/compositing bug. Chromium 152 renders correctly without
it, verified on hardware. /api/status keeps a frozen presentationGpuMode
literal so AV integrations do not break. Native fullscreen is untouched."
```

---

## Gate C — PR 3 verification

- [ ] CI green
- [ ] `grep` for GPU mode identifiers returns exactly one hit (the frozen literal)
- [ ] App launches; a deck presents; transparent/layered slides still render correctly
- [ ] `/api/status` reports `presentationGpuMode: "default"`
- [ ] Settings → Presentation Rendering shows no GPU dropdown, and native fullscreen still saves

---

## Open parameters

Neither blocks this plan:

1. **Version bump size.** This plan uses patch (2.3.11, then 2.3.12) per the repo's patch-only rule. If a 22-major Chromium jump should instead be 2.4.0, change Task 8 Step 1 and the changelog headings.
2. **Blackout window.** Not used in this plan; it constrains *when* PR 2's hardware pass and release are scheduled, and is needed by the companion stay-current plan.
