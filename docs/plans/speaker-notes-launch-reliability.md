# Speaker Notes Launch Reliability — Bugfix Plan

> **Branch:** all work on `bugfix/speaker-notes-launch-reliability` (off `main`).
> **Doc home:** `docs/plans/speaker-notes-launch-reliability.md`.
> **Team:** Three Man Team — Arch = Opus, Bob = Sonnet, Richard = Haiku. Use the cheapest agent that does each phase reliably. Update this doc's status after each phase (per team docs policy).

---

## Context

The web remote (served from `main.js`, port 9595) has two launch buttons, and **both** misbehave:

- **Launch Presentation** → `POST /api/open-presentation` — should open the deck only. **Bug: sometimes opens speaker notes anyway.**
- **Launch with Notes** → `POST /api/open-presentation-with-notes` — should open the deck *and* speaker notes. **Bug: notes sometimes don't appear.**

Reproduces on the dual-monitor setup (presentation + notes on separate screens).

Google Slides has **no API** to open the presenter/notes window. The app fakes an `s` keypress into the presentation `webContents`; Slides reacts by calling `window.open()`, which the app captures via an `app`-level `browser-window-created` listener and adopts as `notesWindow`. That capture (success signal) is reliable. The **trigger** and the **lifecycle** are not.

### Root cause

The auto-notes mechanism uses fire-and-forget timers + an app-level listener bound to **module-global** window references (`presentationWindow`, `notesWindow`), with **no per-launch teardown** and a **blind, unverified trigger**:

1. **"Launch with Notes" misses (notes don't open).** Trigger is a stack of fixed delays (`80ms` focus settle, `250ms` after `did-navigate`, `650ms` after `did-finish-load`) then retries every `700ms` × `8` ≈ **5.6 s total**. It fails when the deck/auth loads slower than that budget, when `s` is sent before Slides' keyboard handler attaches, or when the **two competing trigger paths** (`did-navigate` *and* `did-finish-load`) both press `s` and toggle/duplicate the popup.

2. **"Launch Presentation" leaks notes.** The plain handler never sends `s` — but the with-notes timers (`notesRetryTimer`, the `did-finish-load` `setTimeout`) and the `browser-window-created` listener are **never cancelled** and reference the module globals. Click "Launch with Notes", then click "Launch Presentation" while retries are still pending (≤ ~5.6 s), and a stale retry fires `s` into the **new** plain window → notes open unexpectedly. The stale listener also stays registered on `app`.

### Intended outcome

Each button behaves exactly as labelled, every time: "Launch with Notes" reliably opens notes; "Launch Presentation" never does.

---

## Relevant code (`main.js`)

- Plain launch: `4286-4442` (`/api/open-presentation`).
- With-notes launch: `4444-4655`; fragile trigger at `4534-4611` (`sendSpeakerNotesKey`, `navigationListener`, `did-finish-load`).
- Duplicated notes-capture listener: `4357-4378` and `4514-4532`.
- Reused helpers (keep as-is): `getSpeakerNotesWindowOptions` (`496`), `onNotesWindowCreated` (`779`), `applySpeakerNotesInitialGeometry` (`992`), `attachCrashHandlers`, `toPresentUrl`.
- Other window-open flows to audit for the same leak: IPC/handlers near `2267`, `3550`, `6126`, and the reload/`notesWereOpen` path.

> Line numbers are from this investigation — re-confirm in Phase 1 before editing.

---

## The fix

No public API or UI changes. Constants instead of magic numbers (`NOTES_MAX_ATTEMPTS`, `NOTES_VERIFY_MS`, `NOTES_READY_TIMEOUT_MS`).

### 1. Module-level notes-launch state + teardown
Add near the other window globals:
```js
let notesLaunchToken = 0;              // generation counter; invalidates in-flight loops
let notesRetryTimer = null;            // lifted out of the with-notes handler closure
let notesWindowCreatedListener = null; // lifted out so it can be removed later
```
Add `cancelPendingNotesLaunch()`:
- `notesLaunchToken += 1` (a running loop bails when its captured token no longer matches).
- Clear + null `notesRetryTimer`.
- If set, `app.removeListener('browser-window-created', notesWindowCreatedListener)` and null it.

Call it in the **"Close any existing windows"** block of **both** endpoints (`~4307`, `~4465`) and in the IPC open / reload paths that create a presentation window. **This is the direct fix for bug #2.**

### 2. Shared notes-capture listener
Extract the duplicated listener into `registerNotesWindowListener(notesDisplay)`: stores itself in the module-global `notesWindowCreatedListener`, registers on `app`; on capture sets `notesWindow`, calls `onNotesWindowCreated` + `attachCrashHandlers` + Escape handler + `applySpeakerNotesInitialGeometry`, then removes itself **and** nulls the global. Both endpoints call it (plain launch still wants to adopt a manually-opened notes window).

### 3. Verified, idempotent, cancellable trigger (with-notes only)
Replace `sendSpeakerNotesKey` + `navigationListener` + the `did-finish-load` send with one driver `startNotesLaunchLoop(token, notesDisplay)`:
- Capture `token`; each tick bail if `token !== notesLaunchToken` (superseded), if `presentationWindow` destroyed, or if `notesWindow` already set (done).
- **Wait for present-mode readiness** before the first `s`: poll until `webContents.getURL()` contains `/present/` or `/localpresent` **and** `isLoading() === false`, bounded by `NOTES_READY_TIMEOUT_MS`. Removes reliance on blind delays and the `did-navigate`/`did-finish-load` race.
- `presentationWindow.focus()` + `webContents.focus()`, then send `s` once (keyDown/char/keyUp as today).
- **Wait `NOTES_VERIFY_MS` (~1000 ms)** for the capture listener to set `notesWindow` before deciding to re-press — long enough for Slides to spawn the popup, so we never double-press and toggle it closed.
- Retry until success / supersession / `NOTES_MAX_ATTEMPTS` (~15 ≈ 15 s). **Single driver — no second trigger path.**

Start it once from the with-notes `did-finish-load`, passing the current `notesLaunchToken` (keep the existing `Ctrl+Shift+F5` not-in-present fallback). The plain endpoint **never** starts the loop.

### 4. Keep everything else identical
Display detection, fullscreen chrome, Escape handlers, backup broadcast, and the immediate HTTP response are unchanged.

---

## Phases of work

| Phase | Work | Agent (model) | Why this agent |
|------|------|---------------|----------------|
| **0. Setup** | Create branch `bugfix/speaker-notes-launch-reliability` off `main`; save this doc to `docs/plans/speaker-notes-launch-reliability.md`. | **Richard (Haiku)** | Pure mechanics. |
| **1. Design lock** | Re-confirm root cause against current `main.js`; verify/update the line numbers above; audit IPC + reload paths for the same leak; finalize the change spec & constants. | **Arch (Opus)** | Whole-system reasoning; cheap to get right once. |
| **2. Implement** | Add module state + `cancelPendingNotesLaunch()`; extract `registerNotesWindowListener()`; write `startNotesLaunchLoop()`; wire both endpoints; call teardown in IPC/reload open paths. | **Bob (Sonnet)** | Real coding in a 7700-line file; Sonnet is reliable with this precise spec, cheaper than Opus. |
| **3. Review** | Review the diff for correctness, race-freedom, and no behavior regressions (`/code-review`). | **Arch (Opus)** | Adversarial review is where Opus pays off. |
| **4. Verify** | Richard adds temp `[notes-launch]` logging and runs `yarn start`; **human** runs the dual-monitor test matrix below; Richard removes logging after. | **Richard (Haiku)** + human | Real dual-monitor + live Slides testing is human-driven; agent only scaffolds. |
| **5. Ship** | Bump patch → **v2.3.9** and increment `buildNumber`; update this doc's status; open PR to `main`. | **Richard (Haiku)** | Mechanical release steps. |

Single sequential implementer (Bob) — surgical edits to one file, so no worktree fan-out needed.

---

## Verification (Phase 4 matrix — dual monitor)

1. With temp logging (token, attempt, present-ready, capture), `yarn start`.
2. **Bug #1 fixed:** "Launch with Notes" ~10× across fast and slow/cold-cache loads → notes open **every** time on the notes display; logs show the loop stopping on capture (no excess presses, no toggle-close).
3. **Bug #2 fixed:** Click "Launch with Notes", then within ~2–3 s click "Launch Presentation" → **no** notes window. Repeat several times. A standalone "Launch Presentation" also never opens notes.
4. **No regression:** "Launch Presentation" then manually press `s` → notes still open, captured, and positioned. Reload and IPC open flows behave.
5. **Clean teardown:** no lingering `browser-window-created` listeners/timers across repeated launches (no stray notes windows after cycling).

A focused unit test for `cancelPendingNotesLaunch()` (token increments, timer cleared, listener removed) is worthwhile, but the timing/state nature of the defect makes the manual matrix the primary gate.

---

## Out of scope
- Not unifying the two endpoints into one shared core (large blast radius in a 7700-line file); the extracted helpers remove the worst duplication. Full consolidation is a sensible follow-up.
- Version bump (→ v2.3.9) and `buildNumber` increment happen only at Phase 5, before any packaging build.

## Status log
- _2026-06-22_ — Plan drafted (Arch). Root cause confirmed by code reading. Phase 0 complete: branch `bugfix/speaker-notes-launch-reliability` created off `main`, doc saved here. Next: Phase 1 design lock.
