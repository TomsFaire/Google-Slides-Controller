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
