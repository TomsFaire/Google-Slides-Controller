# plan4.md — API Endpoints (Share Link + QR Control)

**Scope**: Part D only
**Dependencies**: plan3 (getShareLink() and other helpers must exist)
**Parallel**: No — must wait for plan3

## Goal
Add 3 new HTTP API endpoints to the existing server (port 9595) for share link generation and QR overlay control.

## New Endpoints

### POST /api/share-link
```
Body (optional): { forceNew?: boolean }
Response: { ok:true, url, code, expiresAt }

Calls: getShareLink({ forceNew })
Errors:
  - 400 if tunnelPublicUrl not set: { error:"Set Public or tunnel URL first" }
  - 400 if share prefs missing: { error:"Configure shareBaseUrl/shareRegisterUrl/shareApiKey" }
  - 502 if register fails: { error:"Failed to register: {reason}" }
```

### POST /api/show-share-qr
```
Body (optional): { forceNew?: boolean, durationSec?: number }
Response: { ok:true, url, code, expiresAt }

Logic:
  1. Ensure share link exists: link = getShareLink({ forceNew })
  2. Show QR overlay on Presentation display (call showQrOverlay(link, durationSec))
  3. Return share link data

Same errors as /api/share-link above
```

### POST /api/hide-share-qr
```
Body: (none)
Response: { ok:true }

Logic: Hide QR overlay (call hideQrOverlay())
```

## Implementation Notes
- Add to existing HTTP server in main.js (around line 850+)
- Use existing error response patterns
- getShareLink() is called from plan3; ensure main.js is updated first
- showQrOverlay() and hideQrOverlay() are stubbed here (plan5 implements them)
- All endpoints require controller IP allowlist check (existing isControllerAllowedRequest())

## Acceptance Checks
- POST /api/share-link returns share URL
- POST /api/show-share-qr returns share URL
- POST /api/hide-share-qr returns ok:true
- All 3 endpoints return 400 errors for missing config (no crashes)
- Endpoints are callable from CLI/Companion

---

## Token Tracking

**Estimated budget**: 22k tokens (18–28k range)

**Before starting**: Note available tokens in your Claude session.

**During implementation**:
- Read main.js (7700 lines) to locate HTTP server, endpoint patterns — expect ~8–12k tokens
  - Focus on: existing endpoint structure, error response format, IP allowlist checks
  - Less reading than plan3 if main.js patterns were documented in _FINDINGS.md
- Implement 3 endpoints + error handling — ~6–8k tokens
- Test with curl + verify responses — ~2–3k tokens

**Critical checkpoint at 50% through**:
After reading main.js endpoint patterns (~10k tokens used):
1. Check tokens remaining
2. If remaining > 22k tokens → Safe to continue (plan5 needs 28k, but plan5 can be new instance)
3. If remaining < 12k tokens → **PAUSE and report**
   - Commit exploration/stub code
   - Message: "plan4: API endpoint patterns identified, incomplete implementation"
   - Next instance will continue

**Before completing**:
1. Check tokens remaining
2. If tokens remaining < 25k:
   - **Complete plan4 implementation** (you're nearly done)
   - Commit with token count
   - Note: plan5 should start in fresh instance anyway (28k budget)
3. If tokens remaining < 10k:
   - Finish and commit
   - Plan5 will start fresh

**After completion**:
Add this to commit message and/or `_TOKEN_LOG.md`:
```
plan4: ~21k tokens used, X tokens remaining
Findings updated in _FINDINGS.md (API complexity, model recommendations)
Status: READY FOR plan5 & plan6
```

**Important for plan5 decision**:
- In `_FINDINGS.md`, note API integration complexity
- If complexity high → recommend Opus for plan5
- If complexity moderate → Sonnet is fine
- plan5 implementer will check this before starting

**If plan5 is starting in fresh instance**:
- They don't need tokens from your session
- But they MUST read `_FINDINGS.md` for model recommendation
- Make sure `_FINDINGS.md` is committed before they start

---

## Findings Documentation (for plan5 & plan6)

**IMPORTANT**: Before starting, read `_FINDINGS.md` if it exists from plan2 & plan3. Incorporate those notes.

As you implement this plan, **document these findings** in `_FINDINGS.md` to help plan5 & plan6 finalize model choices.

### Plan4 Findings (API Endpoints in main.js)

**Code Patterns Discovered:**
- [ ] API endpoint structure: How are routes defined? (simple if/else, router object, express-like?)
- [ ] Error response format: Is error handling consistent? Can we follow existing pattern?
- [ ] IP allowlist checking: Is `isControllerAllowedRequest()` straightforward to use?
- [ ] Stubbing functions: How do we add placeholders for showQrOverlay/hideQrOverlay?

**Integration Difficulty:**
- [ ] Easy — endpoints fit cleanly into existing pattern, minimal surrounding context needed
  - **Recommendation for plan5**: Sonnet is fine
  - **Recommendation for plan6**: Haiku confirmed
- [ ] Moderate — some context-switching, but manageable
  - **Recommendation for plan5**: Stick with Sonnet
  - **Recommendation for plan6**: Haiku confirmed
- [ ] High — complex integration, lots of interdependencies discovered
  - **Recommendation for plan5**: **Consider Opus for plan5** (window + QR complexity will be compounded)
  - **Recommendation for plan6**: Haiku confirmed (actions don't change)

**Specific Findings for plan5 Implementer (QR Overlay):**
- (What patterns exist for creating windows/overlays in main.js?)
- (Does main.js already have frameless windows or are they new?)
- (What display selection logic exists for presentationDisplayId?)
- (Any existing timeout/auto-hide patterns we can follow?)
- Example: "Window creation is at line 600 with `new BrowserWindow()`. Use similar pattern for QR overlay."

**Specific Findings for plan6 Implementer (Companion Actions):**
- (Confirm that /api/show-share-qr and /api/hide-share-qr exist and return correct JSON)
- (Note the exact response format for debugging in Companion)
- Example: "POST /api/show-share-qr returns { ok:true, url, code, expiresAt }"

---

**When done**:
1. Commit plan4 work
2. Update `_FINDINGS.md` with above checklist + notes
3. Commit findings before plan5 & plan6 start
4. Ensure plan5 reads `_FINDINGS.md` to decide final model (Sonnet vs Opus?)
5. Plan6 can proceed with Haiku confirmed
