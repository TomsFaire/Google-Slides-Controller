# plan3.md — Share Link Generation (Electron main)

**Scope**: Part C only
**Dependencies**: plan2 (settings must exist in prefs)
**Parallel**: No — must wait for plan2

## Goal
Add helper functions to main.js for generating and registering share codes.

## New Functions to Add

### genShareCode()
```
Returns: "word-word-hex8"
- Use 2 word lists (e.g. adjectives + nouns, or random-words npm package)
- Append 8-char hex from crypto.randomBytes(4).toString("hex")
- Example: "happy-cloud-a7f2c19e"
```

### registerShareCode(code)
```
Requires prefs: tunnelPublicUrl, shareBaseUrl, shareRegisterUrl, shareApiKey
POST to shareRegisterUrl with:
  {
    code,
    target: tunnelPublicUrl,
    ttlSec: shareTtlSec,
    key: shareApiKey
  }
Returns: shareUrl ("https://shareBaseUrl/code") or throws error
```

### getShareLink({ forceNew = false })
```
Logic:
  - If cached + not expired and not forceNew:
    - return cached { url, code, expiresAt }
  - Else:
    - code = genShareCode()
    - registerShareCode(code)
    - cache result in prefs
    - return { url, code, expiresAt }
Returns: { url, code, expiresAt } or throws error
```

## Error Handling
- Missing `tunnelPublicUrl` => throw "Set Public or tunnel URL first"
- Missing share prefs => throw "Configure shareBaseUrl/shareRegisterUrl/shareApiKey"
- Register HTTP failure => throw "Failed to register share code: {error}"

## Implementation Notes
- Add to main.js (around line 900+ with other helpers)
- Use native `https.request()` or `node-fetch` for HTTP (check what's already imported)
- Cache in-memory + persist to prefs for durability
- Optional: add simple word lists or use a small dependency

## Acceptance Checks
- genShareCode() returns "word-word-hex8" format
- getShareLink() calls registerShareCode() and caches result
- Error messages are clear for missing config

---

## Token Tracking

**Estimated budget**: 22k tokens (18–28k range)

**Before starting**: Note available tokens in your Claude session.

**During implementation**:
- Read main.js (7700 lines) to understand patterns — expect ~8–12k tokens
  - Focus on: HTTP helpers, preference loading/saving, error patterns
  - You may need multiple passes through this large file
- Understand existing cache patterns — ~2–3k tokens
- Implement 3 helper functions — ~5–7k tokens
- Test + verify caching logic — ~2–3k tokens

**Critical checkpoint at 50% through**:
After reading main.js and understanding patterns (~12k tokens used):
1. Check tokens remaining
2. If remaining > 18k tokens → Safe to continue (plan4 needs ~22k)
3. If remaining < 18k tokens → **PAUSE and report**
   - Commit your exploration/notes
   - Message: "plan3: Main.js explored, insufficient tokens for full implementation"
   - Next instance will continue

**Before completing**:
1. Check tokens remaining
2. If tokens remaining < 22k (enough for plan4):
   - **Finish plan3 implementation**
   - Note tokens used
3. If tokens remaining < 10k:
   - Commit work with complete context
   - Next plan (plan4) must start in fresh instance

**After completion**:
Add this to commit message and/or `_TOKEN_LOG.md`:
```
plan3: ~20k tokens used, X tokens remaining
Findings updated in _FINDINGS.md
Status: READY FOR plan4
```

**If insufficient tokens for plan4**:
- Commit with message including token count
- plan4 implementer will start in new instance
- plan4 should read this commit to understand main.js patterns discovered

---

## Findings Documentation (for plan4 & plan5)

**IMPORTANT**: Before starting, read `_FINDINGS.md` if it exists from plan2. Incorporate plan2's notes.

As you implement this plan, **document these findings** in `_FINDINGS.md` to help plan4 & plan5 determine if model adjustments are needed.

### Plan3 Findings (Electron Helpers in main.js)

**Code Patterns Discovered:**
- [ ] main.js overall complexity: Is it well-modularized or heavily interdependent?
- [ ] Existing HTTP helpers: What pattern does main.js use? (https.request, fetch, axios?)
- [ ] Preference/caching patterns: Is loadPreferences/savePreferences easy to work with?
- [ ] Error handling style: How are errors structured/thrown?
- [ ] Side effects: Are there global state mutations or clean functional patterns?

**Line count observations:**
- [ ] How many lines of code did you add? (aim for 100–150)
- [ ] Did you need to refactor/understand code beyond line 900?
- [ ] How many times did you need to re-read main.js to find patterns?

**main.js Complexity Assessment** (check one):
- [ ] **Clean & organized** — HTTP helpers are clear, caching patterns reusable, state is traceable
  - **Recommendation for plan4**: Haiku might suffice if patterns are this clean
  - **Recommendation for plan5**: Haiku might suffice (window management may differ)
- [ ] **Moderate complexity** — some patterns to follow, needs context but manageable
  - **Recommendation for plan4**: Stick with Sonnet (API endpoint integration is non-trivial)
  - **Recommendation for plan5**: **Stick with Sonnet** (window management + data: URL generation needs power)
- [ ] **High complexity** — tangled dependencies, unclear patterns, lots of context-switching
  - **Recommendation for plan4**: **Use Opus for plan4** (API endpoint integration will be harder)
  - **Recommendation for plan5**: **Use Opus for plan5** (window + QR generation will need careful navigation)

**Specific Findings for plan4 Implementer:**
- (Note HTTP library used in main.js: https.request? fetch?)
- (Note where to add new endpoints for /api/share-link, /api/show-share-qr, /api/hide-share-qr)
- (Note any gotchas with IP allowlist checks or response formatting)
- Example: "HTTP endpoints start at line 850; use https.request pattern, not fetch"

**Specific Findings for plan5 Implementer:**
- (Note display selection patterns if any exist for presentationDisplayId)
- (Note how windows are created and managed in main.js)
- (Note if there are existing overlays or frameless windows we can follow as patterns)
- Example: "Window creation is at line 600; always check for displayId before creating"

---

**When done**:
1. Commit plan3 work
2. Update `_FINDINGS.md` with above checklist + notes
3. Commit findings before plan4 starts
4. Ensure plan4 and plan5 read `_FINDINGS.md` to decide final model choice
