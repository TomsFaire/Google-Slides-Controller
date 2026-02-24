# plan0.md — Master Orchestration (Parallel/Sequential Execution)

Master plan for launching plan1–plan6 across multiple Claude instances with optimal model selection and parallelization.

## Overview

```
STAGE 0 (Check)      Quick verification
                     ↓
STAGE 1 (Parallel)   plan1 (Haiku) + plan2 (Haiku)
                     ↓
STAGE 2 (Sequential) plan3 (Sonnet)
                     ↓
STAGE 3 (Sequential) plan4 (Sonnet)
                     ↓
STAGE 4 (Parallel)   plan5 (Sonnet) + plan6 (Haiku)
                     ↓
STAGE 5 (Verify)     Integration + acceptance tests
```

## Model Strategy

| Model | When to Use | Reason |
|-------|------------|--------|
| **Haiku** | plan1, plan2, plan6 | Small scope, straightforward code, fast + cheap |
| **Sonnet** | plan3, plan4, plan5 | Must read/modify main.js (7700 lines), complex logic |

**Dynamic Model Adjustment**: The model recommendations above are baseline. However:
- **plan2** documents findings about renderer.js complexity
- **plan3** documents findings about main.js complexity + patterns
- **plan4** documents findings that may require upgrading plan5 from Sonnet → **Opus**

See **Findings Workflow** and **Token Management** below.

---

## Token Management & Budget Tracking

**Session Context Limit**: 200,000 tokens per Claude instance
**Safety Margin**: Keep at least 50,000 tokens unused (stop at ~150,000 used)
**Pause Threshold**: If remaining tokens < plan's estimated budget, pause before continuing

### Token Budget Per Plan

| Plan | Model | Scope | Est. Tokens | Range | Notes |
|------|-------|-------|------------|-------|-------|
| **plan1** | Haiku | PHP site (~100 lines) | **5k** | 3–8k | Self-contained, minimal file reading |
| **plan2** | Haiku | Settings UI (~50–100 lines) | **10k** | 8–15k | Reads renderer.js (1300 lines), basic patterns |
| **plan3** | Sonnet | Electron helpers (~100–150 lines) | **22k** | 18–28k | Reads main.js (7700 lines), understanding patterns, caching logic |
| **plan4** | Sonnet | API endpoints (~100–150 lines) | **22k** | 18–28k | Reads main.js again, endpoint patterns, error handling |
| **plan5** | Sonnet/Opus | QR overlay (~150–200 lines) | **28k** (Sonnet) / **20k** (Opus) | 22–35k | Window creation, data: URL, QR generation, npm setup |
| **plan6** | Haiku | Companion actions (~80–120 lines) | **8k** | 5–12k | Small module, simple patterns, straightforward |

### Single-Instance Token Tracking

**Each Claude instance should:**

1. **At start**: Note available tokens (usually 200k limit)
2. **During work**:
   - Implement the assigned plan
   - Keep rough token count (Claude Code displays this)
3. **Before committing**: Check remaining tokens
   - If remaining < next stage's budget → **PAUSE and report**
   - Example: If plan3 uses 20k of 40k budgeted, and plan4 needs 22k, and you have 22k left → OK
   - Example: If plan4 uses 25k and plan5 needs 28k, and you have 23k left → **STOP, insufficient tokens**

4. **Report on completion**:
   ```bash
   # Add to git commit message or _TOKEN_LOG.md:
   # plan3 completed: ~20k tokens used, ~30k remaining
   ```

### Multi-Instance Coordination

**Keep a `_TOKEN_LOG.md` file in repo root:**

```markdown
# Token Usage Log

## Stage 1 (Parallel)
- Instance A (plan1): 5k used, completed ✓
- Instance B (plan2): 10k used, completed ✓

## Stage 2
- Instance C (plan3): 20k used, 20k remaining
  - Status: COMPLETED, safe to proceed

## Stage 3
- Instance D (plan4): 23k used, ? remaining
  - Status: IN PROGRESS
  - Note: If remaining < 22k after completion, plan5 must upgrade to Opus (cheaper tokens)

## Stage 4 (Parallel) — AWAITING TOKEN CHECKS
- Instance E (plan5): Awaiting plan4 token report
- Instance F (plan6): Awaiting plan4 token report
```

**Before launching dependent plans**, check `_TOKEN_LOG.md`:
- If prior stage used more tokens than expected → adjust strategy for next plan
- If tokens are running low → consider:
  - Skipping optional verification steps
  - Using Haiku instead of Sonnet where possible
  - Pausing and starting fresh instance for remaining plans

### Emergency: Running Low on Tokens

**If tokens drop below 50k remaining during a plan:**

1. **Finish current work quickly** (no extra debugging/refactoring)
2. **Commit immediately** with detailed git message:
   ```bash
   git commit -m "plan3: Completed helpers

   Token usage: ~180k of 200k
   Plan complete, safe to proceed to plan4

   Findings for plan4: [summary]"
   ```
3. **Start new instance** for next plan (fresh 200k tokens)
4. **In new instance**, read prior commits to understand context

### Optimizations if Tokens Are Tight

**For Haiku plans (1, 2, 6)**: Use Haiku consistently (cheap, sufficient)
**For Sonnet plans (3, 4, 5)**:
- If tokens tight → Try Haiku first, escalate to Sonnet only if needed
- Provide more specific code snippets to reduce reading overhead
- Skip optional testing/verification steps

---

## Findings Workflow (Dynamic Model Adjustment)

Each plan includes a **"Findings Documentation"** section. Implementers must:

1. **During implementation**, discover code patterns, complexity, and interdependencies
2. **After completion**, update `_FINDINGS.md` in the repo root with:
   - Checklist items (complexity assessment)
   - Specific notes for downstream plans
   - **Model recommendations** for dependent plans

3. **Before next stage starts**, read `_FINDINGS.md` to see if models need adjustment

**Files involved**:
- `_FINDINGS.md` — Central location for cross-plan observations
- `plan2.md` → Updates `_FINDINGS.md` for plan3
- `plan3.md` → Updates `_FINDINGS.md` for plan4 & plan5
- `plan4.md` → Updates `_FINDINGS.md` for plan5 & plan6

**Example escalation**:
- plan3 discovers main.js is heavily interdependent → suggests Opus for plan5
- plan4 discovers API integration is complex → confirms Opus for plan5 needed
- plan5 implementation starts with Opus instead of Sonnet

---

## Stage-by-Stage Execution

### STAGE 0: Pre-Flight Check (10 minutes, 1 instance)

**Purpose**: Verify environment and dependencies before starting.

**Single instance**: Haiku or manual

**Tasks**:
1. Confirm `tunnelPublicUrl` is set in app preferences
2. Confirm BlueHost account access + domain ready for plan1
3. Verify `companion-module-gslide-opener/` exists for plan6
4. Check that main.js and renderer.js exist and are readable
5. Ensure package.json is writable (for adding `qrcode` dep in plan5)

**Output**: Green light to proceed or list blockers

**Command** (if using Claude Code Task):
```bash
claude code plan0:stage0
```

Or manually verify these files exist:
```bash
ls ~/dev/Google-Slides-Controller/{main.js,renderer.js,package.json,preload.js}
ls ~/dev/companion-module-gslide-opener/actions.js 2>/dev/null || echo "Companion module path may differ"
```

**Acceptance**: All files exist, no missing config.

---

### STAGE 1: Parallel — Settings + PHP Site (30–50 minutes, 2 instances)

**Dependencies**: None
**Can run simultaneously**: Yes

#### Instance A: plan1 (PHP Redirect Site)
- **Model**: Haiku
- **Task**: Implement BlueHost PHP site
- **File**: plan1.md
- **Inputs**: BlueHost FTP/domain access
- **Outputs**: Deployed PHP, test redirect working
- **Time**: ~20 minutes

**Launch**:
```bash
# Terminal A
claude code plan1:execute --model haiku
```

Or:
```bash
# Inline prompt
"Implement plan1.md: Tiny PHP redirect site. Deploy to BlueHost if credentials available, or provide complete PHP code ready to deploy. Test and confirm working."
```

#### Instance B: plan2 (Desktop Settings UI)
- **Model**: Haiku
- **Task**: Add share settings to desktop UI
- **File**: plan2.md
- **Inputs**: Existing renderer.js structure
- **Outputs**: Settings UI with 4 new input fields, persisting on restart
- **Time**: ~25 minutes

**Launch**:
```bash
# Terminal B
claude code plan2:execute --model haiku
```

Or:
```bash
# Inline prompt
"Implement plan2.md: Add share settings UI to the desktop app. Read renderer.js and existing HTML, add Share Settings section with 4 inputs (shareBaseUrl, shareRegisterUrl, shareApiKey, shareTtlSec). Ensure they persist via preferences. Make Generate/Show/Hide buttons clickable (wired to IPC stubs is fine). Test UI loads and saves."
```

**Wait for**: Both instances to complete + commit changes
**Check**:
- plan2 UI appears in app
- Settings persist on app restart
- Both report completion ✓

---

### STAGE 2: Sequential — Electron Helpers (25–40 minutes, 1 instance)

**Dependencies**: plan2 must be complete (settings exist)
**Can run in parallel**: No

#### Instance C: plan3 (Share Link Helpers)
- **Model**: Sonnet (needs to read 7700-line main.js)
- **Task**: Add genShareCode(), registerShareCode(), getShareLink() to main.js
- **File**: plan3.md
- **Inputs**: plan2 output (settings merged into repo), main.js source
- **Outputs**: 3 helper functions in main.js, unit-testable
- **Time**: ~30 minutes

**Launch**:
```bash
# Terminal C
claude code plan3:execute --model sonnet
```

Or:
```bash
# Inline prompt
"Implement plan3.md: Add share link helpers to main.js. Read the existing main.js (7700 lines) to understand patterns. Add genShareCode(), registerShareCode(code), and getShareLink({forceNew}). Use existing preferences + HTTP patterns. Test genShareCode() returns valid format. Test getShareLink caches correctly."
```

**Wait for**: Completion + git commit
**After completion**: Implementer must update `_FINDINGS.md` with plan3 findings (code complexity, patterns found, notes for plan4 & plan5)

**Check**:
- genShareCode() returns "word-word-hex8" format ✓
- getShareLink() callable without errors ✓
- Cache logic works ✓
- `_FINDINGS.md` updated with plan3 section ✓

---

### STAGE 3: Sequential — API Endpoints (25–40 minutes, 1 instance)

**Dependencies**: plan3 must be complete (helpers exist)
**Can run in parallel**: No

#### Instance D: plan4 (HTTP API Routes)
- **Model**: Sonnet (modifying main.js, understanding existing API patterns)
- **Task**: Add /api/share-link, /api/show-share-qr, /api/hide-share-qr endpoints
- **File**: plan4.md
- **Inputs**: plan3 output (getShareLink available), main.js HTTP server code
- **Outputs**: 3 new endpoints returning correct JSON, error handling
- **Time**: ~30 minutes

**Launch**:
```bash
# Terminal D
claude code plan4:execute --model sonnet
```

Or:
```bash
# Inline prompt
"Implement plan4.md: Add 3 new API endpoints to main.js HTTP server. Read main.js to understand existing endpoint patterns and IP allowlist checks. Add POST /api/share-link, POST /api/show-share-qr, POST /api/hide-share-qr. Call getShareLink() from plan3. Stub showQrOverlay() and hideQrOverlay() (plan5 will implement). Test with curl; verify error handling for missing config."
```

**Wait for**: Completion + git commit
**After completion**: Implementer must update `_FINDINGS.md` with plan4 findings (API pattern complexity, integration difficulty, potential model upgrade recommendation for plan5)

**Check**:
- `curl -X POST http://localhost:9595/api/share-link` returns valid JSON ✓
- 400 errors for missing config ✓
- showQrOverlay/hideQrOverlay stubs in place ✓
- `_FINDINGS.md` updated with plan4 section ✓

---

### STAGE 4: Parallel — QR Overlay + Companion (35–60 minutes, 2 instances)

**Dependencies**: plan4 must be complete (API endpoints exist)
**Can run simultaneously**: Yes

#### Instance E: plan5 (QR Overlay Window)
- **Model**: Sonnet (default) — **BUT CHECK `_FINDINGS.md` BEFORE LAUNCHING**
  - If plan3 or plan4 noted high complexity in main.js, use **Opus** instead
  - See: `_FINDINGS.md` > "Recommendation for plan5"
- **Task**: Implement QR overlay window on Presentation display
- **File**: plan5.md
- **Inputs**: plan4 output (API endpoints exist), main.js window management patterns, `_FINDINGS.md`
- **Outputs**: showQrOverlay() and hideQrOverlay() functions, frameless window + data: URL
- **Time**: ~35 minutes

**Pre-launch check**:
```bash
cat _FINDINGS.md | grep -A 5 "Recommendation for plan5"
```
If it says "use Opus", adjust launch command below.

**Launch**:
```bash
# Terminal E — DEFAULT (if _FINDINGS.md says Sonnet is OK)
claude code plan5:execute --model sonnet

# OR if _FINDINGS.md recommends Opus:
claude code plan5:execute --model opus
```

Or:
```bash
# Inline prompt
"Implement plan5.md: Add QR overlay window to main.js. Read main.js window creation patterns. Add 'qrcode' npm dependency. Implement showQrOverlay(url, durationSec) and hideQrOverlay(). Create frameless, transparent, always-on-top window with QR image + text fallback. Auto-hide after duration. Test window appears/disappears correctly."
```

#### Instance F: plan6 (Companion Module Actions)
- **Model**: Haiku (small module, simple API calls)
- **Task**: Add show/hide QR actions to Companion module
- **File**: plan6.md
- **Inputs**: plan4 output (API endpoints exist), actions.js file
- **Outputs**: 2 new actions in Companion module
- **Time**: ~15 minutes

**Launch**:
```bash
# Terminal F
claude code plan6:execute --model haiku
```

Or:
```bash
# Inline prompt
"Implement plan6.md: Add 2 actions to companion-module-gslide-opener/actions.js. Add 'Show Share QR' action with durationSec + forceNew options. Add 'Hide Share QR' action. Both call existing HTTP POST helpers to port 9595. Test actions appear in Companion UI and trigger API calls correctly."
```

**Wait for**: Both instances to complete + git commits
**Check**:
- QR window displays on showQrOverlay() call ✓
- QR window auto-hides after timeout ✓
- Companion actions visible in UI ✓
- Companion actions POST to endpoints ✓

---

### STAGE 5: Integration Verification (20–30 minutes, manual + 1 instance)

**Purpose**: End-to-end testing and acceptance.

**Tasks**:
1. **Manual**: Start app, verify settings UI exists
2. **Manual**: Configure share URL + API key in settings
3. **Manual**: Call `/api/share-link` via curl, verify URL returned
4. **Manual**: Call `/api/show-share-qr` via curl, verify QR window appears
5. **Manual**: Scan QR code with phone, verify redirect works
6. **Manual**: Test Companion actions in Companion UI
7. **Automated** (optional): Run any unit tests

**Output**: Green light or list of issues

**Manual test script**:
```bash
# Terminal G - manual verification
cd ~/dev/Google-Slides-Controller
yarn start

# In another terminal:
curl -X POST http://localhost:9595/api/share-link
curl -X POST http://localhost:9595/api/show-share-qr -H "Content-Type: application/json" -d '{"durationSec":10}'
curl -X POST http://localhost:9595/api/hide-share-qr
```

**Acceptance criteria**:
- [ ] Settings UI loads with 4 inputs
- [ ] Share link generation succeeds
- [ ] QR overlay appears on Presentation display
- [ ] QR auto-hides after duration
- [ ] Companion actions trigger without errors
- [ ] Scan QR → redirect to share URL works
- [ ] All 3 API endpoints return valid JSON
- [ ] Error responses are clear (400 for missing config)

---

## Launch Commands Summary

### Option A: Manual (Sequential)
```bash
# Stage 0: Verify (manual or one instance)
# Then...

# Stage 1: Open 2 terminals in parallel
# Terminal A: claude code plan1:execute --model haiku
# Terminal B: claude code plan2:execute --model haiku
# Wait for both...

# Stage 2: Single instance
# Terminal C: claude code plan3:execute --model sonnet
# Wait...

# Stage 3: Single instance
# Terminal D: claude code plan4:execute --model sonnet
# Wait...

# Stage 4: Open 2 terminals in parallel
# Terminal E: claude code plan5:execute --model sonnet
# Terminal F: claude code plan6:execute --model haiku
# Wait for both...

# Stage 5: Manual verification
```

### Option B: Using Claude Code Task Tool (from this instance)
```javascript
// If running from Claude Code, you could spawn agents:
// (pseudo-code, adjust syntax as needed)

await Task.run({ plan: "plan1", model: "haiku" });
await Task.run({ plan: "plan2", model: "haiku" });
// Wait for both...

await Task.run({ plan: "plan3", model: "sonnet" });
// Wait...

await Task.run({ plan: "plan4", model: "sonnet" });
// Wait...

await Task.run({ plan: "plan5", model: "sonnet" });
await Task.run({ plan: "plan6", model: "haiku" });
// Wait for both...
```

---

## Findings Workflow — Cross-Stage Communication

Each plan (2–4) includes code that discovers patterns and complexity. This informs model selection for dependent plans.

### The Loop

1. **plan2 (Settings UI)** completes → adds findings to `_FINDINGS.md`:
   - Complexity of renderer.js
   - Recommendation: Can plan3 use Haiku, or does it need Sonnet?

2. **plan3 (Electron Helpers)** completes → updates `_FINDINGS.md`:
   - Complexity of main.js
   - Interdependencies discovered
   - **Recommendation for plan4 & plan5**: Sonnet OK, or upgrade to Opus?

3. **plan4 (API Endpoints)** completes → updates `_FINDINGS.md`:
   - API integration difficulty
   - **Recommendation for plan5 (QR Overlay)**: Sonnet, or upgrade to Opus?

4. **Before plan5 launches**, check `_FINDINGS.md`:
   ```bash
   cat _FINDINGS.md | grep "Recommendation for plan5"
   ```
   - If it says Opus → use `--model opus`
   - If it says Sonnet OK → use `--model sonnet`

### File Management

**`_FINDINGS.md` structure:**
```markdown
# Findings — Share Link + QR Overlay Implementation

## Plan2 Findings (Settings UI)
- [x] Code complexity: Clean & organized
- [ ] Recommendation for plan3: Haiku/Sonnet/Opus?
- Notes: ...

## Plan3 Findings (Electron Helpers)
- [x] main.js complexity: Moderate
- [ ] Recommendation for plan4: Keep Sonnet
- [ ] Recommendation for plan5: Upgrade to Opus? (window mgmt complex)
- Notes: ...

## Plan4 Findings (API Endpoints)
- [x] API integration: High complexity
- [ ] Recommendation for plan5: CONFIRM Opus needed
- Notes: ...

## Plan6 Findings (Companion Actions)
- [x] Confirmed Haiku works fine
```

**Who updates it:**
- After plan2 completes: plan2 implementer adds "Plan2 Findings" section
- After plan3 completes: plan3 implementer adds "Plan3 Findings" section
- After plan4 completes: plan4 implementer adds "Plan4 Findings" section

**When to read it:**
- Before plan3 launches: Check plan2 recommendation
- Before plan4 launches: Check plan3 recommendation
- Before plan5 launches: Check plan3 & plan4 recommendation → may upgrade from Sonnet to Opus
- Before plan6 launches: Haiku confirmed (no check needed)

---

## Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| Stage 1 blocked | Missing BlueHost access | Skip plan1 deploy, provide PHP code ready to deploy manually |
| Stage 2 fails | plan2 not merged | Check git status, ensure plan2 changes committed |
| Stage 3 fails | main.js unreadable | Verify file exists + readable; share file with instance |
| Stage 4 blocked | API endpoints missing | Re-run plan4 or check for syntax errors |
| QR not appearing | Display ID misconfigured | Verify `prefs.presentationDisplayId` set correctly |
| Companion actions 404 | Wrong port/host | Verify `apiPort` 9595 in settings |

---

## Rollback / Cleanup

If a stage fails:

1. **Identify** which instance/plan failed
2. **Re-run** that plan alone (same model)
3. **Or revert**: `git reset --hard <commit-before-plan>`
4. **Then proceed** with dependent stages

To abort entire orchestration:
```bash
git reset --hard main  # (or main branch)
```

---

## Expected Timeline & Token Usage

### Time + Token Budget Per Stage

| Stage | Plans | Duration | Tokens (est.) | Parallel? | Critical? |
|-------|-------|----------|---------------|-----------|-----------|
| **Stage 0** | — | 10 min | ~1k | Yes | Setup only |
| **Stage 1** | plan1, plan2 | 50 min | 5k + 10k = **15k** | Yes | Baseline setup |
| **Stage 2** | plan3 | 40 min | **22k** | No | Heaviest reading (main.js) |
| **Stage 3** | plan4 | 40 min | **22k** | No | Heaviest reading (main.js) |
| **Stage 4** | plan5, plan6 | 60 min | 28k + 8k = **36k** | Yes | plan5 may need Opus |
| **Stage 5** | — | 30 min | ~2k | Yes | Manual verification |
| **Total** | 1–6 | **3–4 hrs** | **~97k tokens** | — | — |

### Cumulative Token Usage per Instance

If running each plan in a single long-lived instance (not recommended):
```
After Stage 1: 15k tokens used, 185k remaining
After Stage 2: 15k + 22k = 37k, 163k remaining
After Stage 3: 37k + 22k = 59k, 141k remaining ← Safe buffer
After Stage 4: 59k + 36k = 95k, 105k remaining ← Getting tight
After Stage 5: 95k + 2k = 97k, 103k remaining ← Below 50k margin
```

**Recommendation**: Run plans 2–6 in separate instances to avoid token exhaustion:
- Instance A: plan1 + plan2 (15k tokens, plenty of room)
- Instance B: plan3 (22k tokens, then start fresh)
- Instance C: plan4 (22k tokens, then start fresh)
- Instance D: plan5 (28k+ tokens if Opus, then start fresh)
- Instance E: plan6 (8k tokens)

### Per-Instance Token Checkpoints

**Each plan must check tokens:**

| Plan | Checkpoint at 50% | Action if Tokens Low | Recovery |
|------|-------------------|----------------------|----------|
| plan1 | After design (~2k) | Finish quickly | Can pause/resume |
| plan2 | After reading renderer (~4k) | Finish (10k cheap) | Can pause/resume |
| plan3 | After reading main.js (~12k) | **PAUSE if <10k left** | Start fresh instance |
| plan4 | After reading main.js (~10k) | **PAUSE if <12k left** | Start fresh instance |
| plan5 | After reading main.js (~12k) | **PAUSE if <15k left** | Start fresh instance |
| plan6 | After patterns (~3k) | Finish quickly | Almost done |

### Token Safety Rules

1. **Never start a plan if tokens < (estimated budget + 10k buffer)**
   - Example: Don't start plan3 (22k) if you have < 32k tokens
2. **At 50% through a plan, check tokens:**
   - If remaining < plan's budget → pause and report
   - If remaining >= budget → safe to finish
3. **Before dependent plan starts**, read prior commits:
   - Check `_TOKEN_LOG.md` for actual tokens used
   - If lower than estimate → Good, you have buffer
   - If higher → Adjust expectations for next plan
4. **Emergency**: If you hit critical low (<30k remaining):
   - Finish current work immediately
   - Commit with full context
   - Next plan uses fresh instance

### Optimizing Token Usage

- **Haiku for plans 1, 2, 6**: Cheaper, sufficient, saves ~5k tokens combined
- **Sonnet vs Opus for plan3–5**: Use _FINDINGS.md to determine; Opus may be necessary but costs more
- **Batch reading**: Have instance read multiple files at once (main.js + package.json) instead of separately
- **Targeted focus**: Skip reading entire files; ask for specific line ranges or patterns
- **Skip optional testing**: For time-critical plans, skip manual verification; rely on linting

---

## _TOKEN_LOG.md Template

Create and maintain this file in the repo root to track token usage across all plans:

```markdown
# Token Usage Log — Share Link + QR Overlay Implementation

## Summary
- Total tokens budgeted: ~97k
- Total tokens used: [TBD after completion]
- Efficiency: [TBD]

## Stage 1 (Parallel)

### Instance A — plan1 (PHP Redirect Site)
- Model: Haiku
- Duration: ~20 min
- Tokens used: 4k
- Tokens remaining: [N/A — fresh instance]
- Status: ✓ COMPLETED
- Notes: BlueHost PHP deployed and tested

### Instance B — plan2 (Settings UI)
- Model: Haiku
- Duration: ~25 min
- Tokens used: 10k
- Tokens remaining: [N/A — fresh instance]
- Status: ✓ COMPLETED
- Notes: Settings UI added, all fields persist

## Stage 2 (Sequential)

### Instance C — plan3 (Electron Helpers)
- Model: Sonnet
- Duration: ~30 min
- Tokens used: 20k
- Tokens remaining after completion: 20k
- Status: ✓ COMPLETED
- Token checkpoint at 50%: 12k used, safe to continue
- Notes: main.js patterns clean, getShareLink caching works
- Findings: [ref _FINDINGS.md for plan4/5 recommendations]

## Stage 3 (Sequential)

### Instance D — plan4 (API Endpoints)
- Model: Sonnet
- Duration: ~35 min
- Tokens used: 21k
- Tokens remaining after completion: ?k (depends on Instance C handoff)
- Status: ✓ COMPLETED
- Token checkpoint at 50%: 10k used, safe to continue
- Notes: 3 endpoints working, error handling verified
- Findings: [ref _FINDINGS.md for plan5 model recommendation]

## Stage 4 (Parallel)

### Instance E — plan5 (QR Overlay)
- Model: [Sonnet/Opus — check _FINDINGS.md]
- Duration: ~40 min
- Tokens used: 25k
- Status: ✓ COMPLETED
- Token checkpoint at 50%: 12k used, safe to continue
- Notes: QR window working, auto-hide verified

### Instance F — plan6 (Companion Actions)
- Model: Haiku
- Duration: ~20 min
- Tokens used: 7k
- Status: ✓ COMPLETED
- Notes: 2 actions added, Companion UI verified

## Stage 5 (Verification)

### Manual Integration Tests
- Duration: ~30 min
- Status: ✓ PASSED
- Notes: Full end-to-end test completed successfully

## Recommendations for Next Iteration

- [Any model changes recommended?]
- [Any process improvements?]
- [Actual vs estimated token usage analysis]
```

**How to use this file**:

1. **Create** `_TOKEN_LOG.md` in repo root before starting
2. **After each plan completes**, update the relevant section:
   ```bash
   git add _TOKEN_LOG.md
   git commit -m "plan2: Completed — 10k tokens used

   Token summary for stage: 10k used, 190k remaining"
   ```
3. **Before each dependent plan**, read `_TOKEN_LOG.md`:
   ```bash
   cat _TOKEN_LOG.md | grep -A 5 "Instance [NEXT]"
   ```
4. **At end of orchestration**, review total tokens used vs budgeted

---

## Notes

- **Stages 1 & 4 can save ~50 min** with true parallelization
- **Each plan should run in its own instance** to avoid token exhaustion
- Use **Haiku** for stages 1, 2, 6 to reduce costs
- Use **Sonnet for stages 3, 4** because main.js (7700 lines) is expensive to read/modify
- **Upgrade to Opus if needed** for plan5 QR window (check _FINDINGS.md for complexity assessment)
- **Token tracking is mandatory** — each implementer must report tokens in commit messages
- **Plan0 itself does not write code** — it orchestrates the execution of plan1–plan6
