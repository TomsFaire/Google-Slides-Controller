# DeckLink backend seam: human validation

Branch `feature/native-decklink-addon`, stacked on `feature/decklink-output` (PR #23). It changes **how** the app talks to the card, not what it does. This checklist was run on hardware on 2026-10-23, see [Results](#results-2026-10-23). **The PR targets `feature/decklink-output`, not `main`.**

## What changed (so you know what to watch)

| Commit | Change | Risk on hardware |
|---|---|---|
| `9f653ed` | All macadam calls moved behind a backend interface (`src/decklink-backend.js`, `src/backends/macadam-backend.js`); `macadam-worker.js` is now `decklink-worker.js`; `MacadamProvider` is `WorkerProvider` | Should be invisible. Any regression in start/stop/frames points here. |
| `85a9436` | Pixel format is negotiated: BGRA is tried first, YUV only if the device rejects the mode. The hard-coded `yuv` flags are gone | 1080p50/59.94/60 now make **two** attempts. If those modes regress, look here. |
| `19a9049` | The Format dropdowns are filled from what each device reports | Mode names are matched from macadam's output; unconfirmed against a real device. |

Also in the branch from the base: FFmpeg fallback removed, DeckLink API version shown in the tab, resize quality `good`, per-stage worker timings in the stats line (those base changes are also untested on hardware).

## Setup

```bash
cd /Users/tom/Documents/gslide-opener/.claude/worktrees/native-decklink
export NVM_DIR="$HOME/.nvm" && source "$NVM_DIR/nvm.sh"
pkill -f decklink-worker
yarn test        # expect 160 passing
```

This worktree uses the other worktree's `node_modules` (a symlink), so the macadam build is the one that already works on your machine. Quit any other copy of the app first.

## Checks

Tick each one. Paste anything that fails or surprises you.

### 1. Hardware regression (backend seam + pixel format)

- [ ] 1080p60
  ```bash
  DECKLINK_MODE=1080p60 yarn test:decklink-hw
  ```
  Expect **7/7**, bars on the Video Assist, and a log line `output started: 1080p60, pixel format yuv`.
- [ ] 1080p59.94: `DECKLINK_MODE=1080p5994 yarn test:decklink-hw` → 7/7, `pixel format yuv`.
- [ ] 1080p29.97: `DECKLINK_MODE=1080p2997 yarn test:decklink-hw` → 7/7, `pixel format bgra`.
- [ ] Colour still right: `yarn decklink:color-chart` looks the same as before on the Video Assist. Optional A/B: `DECKLINK_PIXEL_FORMAT=yuv` and `=bgra` on a mode that allows both.

### 2. Device and mode list

- [ ] `yarn decklink:info`. **Paste the whole output** (I need the real mode names to check the matcher).
- [ ] `yarn decklink:probe-modes` runs through the modes. Note which pass.

### 3. The app

```bash
yarn start 2>&1 | grep --line-buffered "\[decklink\]"
```

- [ ] DeckLink tab, Provider line reads **`Provider: macadam (DeckLink API 16.4.0)`**.
- [ ] The Device dropdown lists the UltraStudio by name.
- [ ] The Format dropdown lists the modes the 3G supports and no UHD/10-bit extras. Expected: 1080p 59.94 / 60 / 50 / 30 / 29.97 / 25, 720p 59.94 / 50. Interlaced modes are intentionally not offered.
- [ ] Enable slides output, pick 1080p29.97, **Save**: the light goes green, your slides appear on the output, advancing slides follows without lag.
- [ ] Switch to 1080p59.94 or 60, **Save**: output comes back (YUV path). The log says `pixel format yuv`.
- [ ] Toggle the output off, Save, on, Save: restarts cleanly, no "hardware currently active" error.
- [ ] Every 5 s the log prints a `[decklink] slides:` line. **Paste 3 or 4 of them**, including the new `worker … round-trip … (yuv …, displayFrame …)` part.

### 4. Unhappy paths

- [ ] USB4 cable unplugged, app started: the tab still opens and shows the full static Format list, no crash.
- [ ] Optional: edit `displayMode` in `preferences.json` to a mode the card does not offer (for example `720p5994` is fine, pick one you removed from the list), reopen the tab: it switches to a supported format and shows the explanation under the dropdown.

## Reporting back

Paste: the `decklink:info` output, the stats lines, and anything that failed. If a hardware check in section 1 fails, say which mode, and I will look at `85a9436` first.

## When it passes

1. Open a PR from `feature/native-decklink-addon` into **`feature/decklink-output`**.
2. That branch (PR #23) still has its own unchecked tests; this one merges into it, and the whole DeckLink set rolls into `main` from there.
3. If PR #23 changes after more testing, this branch is rebased or merged onto it before the PR.

## Results (2026-10-23)

UltraStudio Express Monitor 3G, Desktop Video 16.4.0, macOS, UHD displays. macadam backend throughout.

| Check | Result |
|---|---|
| 1. 1080p60, 1080p59.94 (`yuv`), 1080p29.97 (`bgra`) hardware test | 7/7 each, colour bars correct on the Video Assist |
| 1. Colour chart | Colours looked right at several resolutions |
| 2. `decklink:info`, `decklink:probe-modes` | All 9 modes opened (1080i59.94 has since been removed from the app) |
| 3. App with real slides | Slides appeared and looked good, next slide and animated GIFs fast enough |
| 3. Mode switch to 59.94, off/on restart, Format dropdown contents | Not checked |
| 4. Unhappy paths (cable unplugged) | Not checked |

Found on the way: the probe script and hardware test sent black frames because their fake window had no `getSize()`, while still passing. Fixed in `bf13433`. Stats lines from the real-slides run are in the baseline table in [DECKLINK.md](DECKLINK.md#measured-baseline-macadam-backend).
