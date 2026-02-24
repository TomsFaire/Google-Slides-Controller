# plan6.md — Companion Module Actions

**Scope**: Part F only
**Dependencies**: plan4 (API endpoints must exist at /api/show-share-qr and /api/hide-share-qr)
**Parallel**: Can run in parallel with plan5

## Goal
Add 2 new action definitions to the Companion module (`companion-module-gslide-opener/actions.js`) for QR control.

## New Actions

### "Show Share QR"
```
Action ID: show_share_qr
Options:
  - durationSec: number (default: 20, range 5–300)
  - forceNew: boolean (default: false)

Behavior:
  POST http://{host}:{port}/api/show-share-qr
  Body: { durationSec, forceNew }

  Feedback (optional):
    - onSuccess: display "QR shown for X sec"
    - onError: display error message
```

### "Hide Share QR"
```
Action ID: hide_share_qr
Options: (none)

Behavior:
  POST http://{host}:{port}/api/hide-share-qr

  Feedback (optional):
    - onSuccess: display "QR hidden"
    - onError: display error message
```

## Implementation Notes
- Add to existing actions.js structure (follow existing action patterns)
- Use existing HTTP POST helper already in the module
- Reuse existing host/port config from module settings
- Both actions should include error handling (bad response, network timeout)
- Optional: add feedback labels for Companion UI display

## File Structure
- `companion-module-gslide-opener/actions.js` - add 2 action defs
- No new files required

## Acceptance Checks
- Both actions appear in Companion UI action list
- "Show Share QR" option dropdowns work
- "Hide Share QR" button triggers without options
- Successful POST returns 200 response
- Failed POST displays error message in Companion log

---

## Token Tracking

**Estimated budget**: 8k tokens (5–12k range)

**Before starting**: Note available tokens in your Claude session.

**During implementation**:
- Read actions.js to understand existing action patterns — expect ~2–3k tokens
- Understand HTTP POST helper in module — ~1–2k tokens
- Implement 2 actions + options — ~3–4k tokens
- Test actions trigger correctly — ~1–2k tokens

**Token checkpoint at 50% through**:
After reading actions.js patterns (~3k tokens used):
1. Check tokens remaining
2. If remaining > 8k → Safe to continue
3. If remaining < 5k → This is very unusual, but:
   - Finish plan6 quickly (only 5k left to do)
   - Commit immediately
   - You're still OK

**Before completing**:
1. Check tokens remaining
2. If tokens remaining < 8k:
   - **Finish plan6 implementation** (small scope, nearly done)
   - Commit with token count
3. This is the last plan, so token shortage is not a blocker

**After completion**:
Add this to commit message and/or `_TOKEN_LOG.md`:
```
plan6: ~7k tokens used, X tokens remaining
Companion actions implemented and tested
Status: COMPLETE — Ready for integration verification (Stage 5)
```

**After all plans complete**:
Verify commit includes:
- All plan implementations merged to main
- `_TOKEN_LOG.md` updated with final token counts
- `_FINDINGS.md` documented across all stages
- Ready for manual integration testing in Stage 5
