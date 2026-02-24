# plan2.md — App Settings UI (Desktop)

**Scope**: Part B only
**Dependencies**: None (can run in parallel)
**Can run in parallel**: Yes

## Goal
Add user-facing settings in the desktop UI for share link configuration.

## New Preference Fields
Add to preferences JSON structure:
```
{
  shareBaseUrl: "https://slides.example.com",
  shareRegisterUrl: "https://slides.example.com/api/register",
  shareApiKey: "secret-key-from-bluehost",
  shareTtlSec: 86400,
  shareLastCode: "word-word-abcd1234",        // optional cache
  shareLastUrl: "https://slides.example.com/word-word-abcd1234",
  shareLastExpiresAt: 1708972800
}
```

## UI Changes (renderer.js / HTML)
Add a new "Share Settings" section with:
- Input: `shareBaseUrl`
- Input: `shareRegisterUrl`
- Input: `shareApiKey` (password field, masked)
- Input: `shareTtlSec` (number, default 86400)
- Display: `shareLastUrl` (read-only, for reference)
- Buttons:
  - "Generate Share Link"
  - "Show QR on Presentation"
  - "Hide QR"

## Implementation Notes
- Add form inputs to existing settings UI
- Wire up load/save via existing `loadPreferences()`/`savePreferences()` pattern
- Validate URLs before saving
- All IPC calls to main for actual share link generation/display (main.js handles)

## Acceptance Checks
- Settings UI displays all 4 input fields
- Inputs persist across app restart
- Generate/Show/Hide buttons are clickable (wired to IPC, no-op is fine at this stage)

---

## Token Tracking

**Estimated budget**: 10k tokens (8–15k range)

**Before starting**: Note available tokens in your Claude session.

**During implementation**:
- Read renderer.js (1300 lines) — expect ~3–5k tokens
- Understand preferences pattern — ~2–3k tokens
- Implement UI form + IPC wiring — ~3–5k tokens
- Test + verify — ~1–2k tokens

**Before completing**:
1. Check tokens remaining in your session
2. If tokens remaining < 18k (enough for plan3):
   - **Continue to completion** (plan3 doesn't depend on plan2's token usage)
3. If tokens remaining < 18k but also < 8k:
   - **Finish this plan** (you're almost done)
   - Commit work with message: `plan2: UI implemented, X tokens used`
   - Note the token count

**After completion**:
Add this line to your commit message or to `_TOKEN_LOG.md`:
```
plan2: ~10k tokens used, X tokens remaining
Status: READY FOR plan3
```

**If running low on tokens** (< 50k remaining):
- Finish plan2 quickly
- Commit with full context notes
- Next plan (plan3) will start in fresh instance

---

## Findings Documentation (for plan3)

As you implement this plan, **document these findings** to help plan3 decide if Sonnet is sufficient or if it should use a more capable model.

**Create or update**: `_FINDINGS.md` in repo root with this section:

### Plan2 Findings (Settings UI)

**Code Patterns Discovered:**
- [ ] Is `loadPreferences()`/`savePreferences()` simple JSON I/O or complex?
- [ ] How are IPC calls structured? (simple `ipc.send()` or complex messaging?)
- [ ] Renderer size/complexity: Is renderer.js well-organized or tangled?
- [ ] Are there existing settings sections we can closely follow as templates?

**Findings Summary** (check one):
- [ ] **Simple patterns found** — renderer.js is well-organized, IPC is straightforward
  - **Recommendation for plan3**: Haiku may be sufficient for main.js if it's equally clean
- [ ] **Moderate complexity** — some patterns to follow, but code is readable
  - **Recommendation for plan3**: Stick with Sonnet (main.js is 7700 lines, need power to navigate it)
- [ ] **High complexity** — tangled code, unclear patterns, lots of state management
  - **Recommendation for plan3**: **Use Opus for plan3** (main.js likely equally complex)

**Notes for plan3 implementer**:
- (Write any specific insights about renderer.js patterns, IPC structure, or gotchas)
- Example: "Preferences are stored at `path.join(app.getPath('userData'), 'preferences.json')`"

---

**When done**: Commit plan2 work, then update `_FINDINGS.md` and commit again before starting plan3
