# Orchestration Summary — Token-Aware Multi-Instance Planning

This document summarizes all the changes made to support parallel/sequential execution with token tracking and dynamic model selection.

## What Changed

### New Files Created

1. **plan0.md** — Master orchestration plan
   - 5 stages with clear dependencies
   - Per-stage model selection (Haiku/Sonnet/Opus)
   - Detailed token budgeting
   - Token checkpoint system
   - _TOKEN_LOG.md format + usage

2. **plan1.md** — PHP Redirect Site (updated)
   - Added "Token Tracking" section
   - Checkpoint at 50%
   - Commit message format

3. **plan2.md** — Settings UI (updated)
   - Added "Findings Documentation" section
   - Documents renderer.js complexity
   - Recommends model for plan3
   - Added "Token Tracking" section

4. **plan3.md** — Electron Helpers (updated)
   - Added "Findings Documentation" section
   - Documents main.js complexity + patterns
   - Recommends models for plan4 & plan5
   - Added "Token Tracking" section with critical checkpoint at 50%

5. **plan4.md** — API Endpoints (updated)
   - Added "Findings Documentation" section
   - Documents API integration difficulty
   - Confirms/escalates plan5 model choice
   - Added "Token Tracking" section

6. **plan5.md** — QR Overlay (updated)
   - Added dynamic model selection (check _FINDINGS.md before launch)
   - Added "Token Tracking" section
   - Can use Sonnet or Opus based on findings

7. **plan6.md** — Companion Actions (updated)
   - Added "Token Tracking" section
   - Simplest plan, last in pipeline

8. **TOKEN_TRACKING_CHECKLIST.md** — Quick reference for implementers
   - Checkpoint checklist (25%, 50%, 75%, final)
   - Commit message format
   - Emergency procedures
   - Budget quick reference
   - Optimization tips

9. **ORCHESTRATION_SUMMARY.md** — This file

## Key Features

### Token Management

**Per-plan budgets:**
- plan1: 5k tokens (Haiku)
- plan2: 10k tokens (Haiku)
- plan3: 22k tokens (Sonnet)
- plan4: 22k tokens (Sonnet)
- plan5: 28k tokens (Sonnet) or 20k (Opus)
- plan6: 8k tokens (Haiku)
- **Total: ~97k tokens** for full orchestration

**Safety margins:**
- 200k token limit per session
- Keep 50k tokens unused (stop at ~150k used)
- Critical checkpoint at 50% completion
- Pause if insufficient tokens for next stage

### Dynamic Model Selection

**Baseline:**
- Haiku for plans 1, 2, 6 (small scope, cheap)
- Sonnet for plans 3, 4, 5 (large files, complex logic)

**Dynamic adjustment:**
- plan2 documents renderer.js complexity → suggests model for plan3
- plan3 documents main.js complexity → suggests models for plan4 & plan5
- plan4 documents API complexity → confirms or escalates plan5 model
- plan5 checks _FINDINGS.md before launch → may upgrade Sonnet → Opus

### Findings Workflow

Each plan (2-4) discovers code patterns and complexity:
- plan2 → documents renderer.js, recommends for plan3
- plan3 → documents main.js, recommends for plan4 & plan5
- plan4 → documents API integration, final recommendation for plan5

All findings stored in central `_FINDINGS.md` file.

### Multi-Instance Coordination

**Recommended execution:**
- Instance A: plan1 + plan2 (15k tokens, parallel)
- Instance B: plan3 (22k tokens, starts after plan2)
- Instance C: plan4 (22k tokens, starts after plan3)
- Instance D: plan5 (28k+ tokens, starts after plan4)
- Instance E: plan6 (8k tokens, starts after plan4, parallel with D)

**Token handoff via git:**
- Each instance commits with token count
- Central `_TOKEN_LOG.md` tracks all instances
- Next stage reads token usage before starting

### Checkpoints

**At 50% completion (CRITICAL):**
- Check tokens remaining
- If < plan's budget → pause, commit, use fresh instance
- If >= budget → safe to continue

**Before final commit:**
- Report exact tokens used
- Include in commit message
- Update _TOKEN_LOG.md

## Usage Workflow

### Before Starting

1. Create `_TOKEN_LOG.md` in repo root
2. Ensure all plan files (plan0-6) are in place
3. Verify TOKEN_TRACKING_CHECKLIST.md is available

### Stage 1: Parallel (plan1 + plan2)

**Instance A** (Haiku):
```bash
# Start plan1
# At 50%: Check tokens (should have ~195k remaining)
# Commit with token count
```

**Instance B** (Haiku):
```bash
# Start plan2
# At 50%: Check tokens (should have ~190k remaining)
# Commit with token count
# Update _FINDINGS.md with renderer.js findings
```

**Wait for both to complete** before stage 2.

### Stage 2: Sequential (plan3)

**Before launching:**
```bash
cat _FINDINGS.md | grep "Plan2 Findings"
cat _TOKEN_LOG.md
```

**Instance C** (Sonnet):
```bash
# Start plan3
# Checkpoint 1 (after reading main.js ~12k): Verify tokens
# Checkpoint 2 (50% complete ~12k): CRITICAL — must have >=18k
# Complete and commit with token count
# Update _FINDINGS.md with main.js findings + plan4/5 recommendations
```

### Stage 3: Sequential (plan4)

**Before launching:**
```bash
cat _FINDINGS.md | grep "Plan3 Findings"
cat _TOKEN_LOG.md
```

**Instance D** (Sonnet):
```bash
# Start plan4
# Checkpoint at 50%: Verify tokens
# Complete and commit with token count
# Update _FINDINGS.md with API complexity + plan5 final recommendation
```

### Stage 4: Parallel (plan5 + plan6)

**Before launching:**
```bash
# CRITICAL: Check plan5 model recommendation
cat _FINDINGS.md | grep "Recommendation for plan5"
# If Opus → use --model opus
# If Sonnet → use --model sonnet
```

**Instance E** (Sonnet or Opus based on _FINDINGS.md):
```bash
# Start plan5 with correct model
# Checkpoint at 50%: Verify tokens
# Complete and commit
```

**Instance F** (Haiku):
```bash
# Start plan6 (independent)
# Complete and commit
```

### Stage 5: Verification (manual)

- Verify all commits merged
- Run integration tests
- Confirm QR works end-to-end

## File References

**Updated files:**
- `plan0.md` — Master orchestration (NEW)
- `plan1.md` — Token tracking added
- `plan2.md` — Token tracking + findings added
- `plan3.md` — Token tracking + findings added
- `plan4.md` — Token tracking + findings added
- `plan5.md` — Dynamic model selection + token tracking added
- `plan6.md` — Token tracking added
- `PLAN_INDEX.md` — Existing, no changes needed
- `plan.md` — Original detailed plan, kept for reference

**New files:**
- `TOKEN_TRACKING_CHECKLIST.md` — Quick reference for implementers
- `ORCHESTRATION_SUMMARY.md` — This file
- `_TOKEN_LOG.md` — Created by implementers, tracks actual usage
- `_FINDINGS.md` — Created during implementation, documents complexity

## Commit Message Format

Each plan should commit with:

```
plan[N]: [Status — Implementation complete/Paused/etc]

Model used: [Haiku/Sonnet/Opus]
Tokens used: ~[X]k tokens
Tokens remaining: ~[Y]k tokens
Status: [COMPLETED/PAUSED]

[Plan-specific notes]

[If applicable] Findings updated: YES
```

## Emergency Recovery

**If a plan runs out of tokens mid-way:**

1. Commit immediately with detailed notes
2. Mark in _TOKEN_LOG.md as PAUSED
3. Start fresh instance for continuation
4. New instance reads prior commits for context
5. No loss of progress, just continuation

**If dependent plan can't start:**

1. Check _TOKEN_LOG.md for prior stage token usage
2. If prior stage used too many tokens → use fresh instance
3. All plans can run independently if needed

## Success Criteria

✓ All 6 plans completed
✓ Total tokens used ~97k (well under 200k limit)
✓ _TOKEN_LOG.md fully populated
✓ _FINDINGS.md fully populated with complexity assessments
✓ Dynamic model selection used for plan5 based on findings
✓ All code committed with token tracking
✓ Integration verification passed

## Expected Timeline

- Stage 0: 10 min (setup)
- Stage 1: 50 min (parallel, 2 instances)
- Stage 2: 40 min (plan3, wait for stage 1)
- Stage 3: 40 min (plan4, wait for stage 2)
- Stage 4: 60 min (parallel, 2 instances, wait for stage 3)
- Stage 5: 30 min (manual verification)
- **Total: 2.5–3.5 hours**

With proper parallelization: **~2.5 hours**

## Next Steps

1. Review plan0.md for complete orchestration strategy
2. Review TOKEN_TRACKING_CHECKLIST.md before starting each plan
3. Create _TOKEN_LOG.md in repo root
4. Start Stage 1 with instances A (plan1) and B (plan2)
5. Track tokens at each checkpoint
6. Update _FINDINGS.md as you discover complexity
7. Proceed to dependent stages as each completes

---

**All changes support the goal**: Execute 6 plans across multiple Claude instances in parallel where possible, tracking tokens at each step, and using findings to dynamically select models, all while maintaining a safety buffer of unused tokens.
