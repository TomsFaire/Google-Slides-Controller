# Launch Checklist — 10pm PST (Fresh Session, 200k Tokens)

When your session resets at 10pm PST, use this checklist to kick off the full orchestration.

## Pre-Launch Verification (Before 10pm)

- [ ] All plan files are committed to git:
  - [ ] plan0.md (master orchestration)
  - [ ] plan1.md–plan6.md (all with token tracking)
  - [ ] TOKEN_TRACKING_CHECKLIST.md
  - [ ] ORCHESTRATION_SUMMARY.md
  - [ ] PLAN_INDEX.md
  - [ ] plan.md (original detailed plan, reference)

- [ ] Git status is clean:
  ```bash
  cd ~/dev/Google-Slides-Controller
  git status
  # Should show "nothing to commit, working tree clean"
  ```

- [ ] You have access to:
  - [ ] BlueHost FTP credentials (for plan1 PHP deployment)
  - [ ] Google Slides Controller app directory
  - [ ] Companion module directory (for plan6)

---

## At 10pm PST — Session Reset (200k tokens available)

### 1. Verify Fresh Session

```bash
# Check that you're starting fresh
echo "Session started at: $(date)"
# Should show current time around 10pm PST
```

### 2. Create Token Log File

```bash
cd ~/dev/Google-Slides-Controller

# Create _TOKEN_LOG.md
cat > _TOKEN_LOG.md << 'EOF'
# Token Usage Log — Share Link + QR Overlay Implementation

**Session Start**: [Time at 10pm PST]
**Initial Budget**: 200,000 tokens
**Safety Margin**: Stop at 150,000 tokens used (50k reserve)

## Stage 1 (Parallel) — plan1 + plan2

### Instance A — plan1 (PHP Redirect Site)
- Model: Haiku
- Status: AWAITING LAUNCH
- Tokens used: [TBD]
- Tokens remaining: [TBD]

### Instance B — plan2 (Settings UI)
- Model: Haiku
- Status: AWAITING LAUNCH
- Tokens used: [TBD]
- Tokens remaining: [TBD]

## [Remaining stages will be updated as plans complete]
EOF

git add _TOKEN_LOG.md
git commit -m "Initial: Token log for 10pm PST orchestration session"
```

### 3. Open Terminal Windows

Open **2 terminal windows** for Stage 1 (parallel execution):

**Terminal A** — Plan 1 (PHP Site):
```bash
# In Terminal A:
cd ~/dev/Google-Slides-Controller
# Keep this window ready for plan1 launch

echo "Terminal A: Ready for plan1 (PHP Redirect Site)"
echo "Estimated duration: ~20 min, ~5k tokens"
```

**Terminal B** — Plan 2 (Settings UI):
```bash
# In Terminal B:
cd ~/dev/Google-Slides-Controller
# Keep this window ready for plan2 launch

echo "Terminal B: Ready for plan2 (Settings UI)"
echo "Estimated duration: ~25 min, ~10k tokens"
```

### 4. Launch Stage 1 (Parallel)

**In Terminal A** — Start plan1:
```bash
# Read plan1.md
cat plan1.md

# Then in Claude Code:
# Prompt: "Implement plan1.md: Tiny PHP redirect site on BlueHost..."
# Model: Haiku
```

**In Terminal B** — Start plan2:
```bash
# Read plan2.md
cat plan2.md

# Then in Claude Code:
# Prompt: "Implement plan2.md: Add share settings to desktop UI..."
# Model: Haiku
```

**Expected time**: Both complete in ~50 min (parallel)

### 5. After Stage 1 Completes

```bash
# Both implementers commit their work with token counts:
# git commit -m "plan1: Completed — 5k tokens used, 195k remaining"
# git commit -m "plan2: Completed — 10k tokens used, 190k remaining"

# Update _TOKEN_LOG.md with actual numbers
git add _TOKEN_LOG.md
git commit -m "Stage 1 complete: 15k tokens used, 185k remaining"

# Verify both are merged
git log --oneline | head -5
```

---

## Stage 2 — Plan 3 (Sequential)

**Start after**: Stage 1 complete (~10:50pm PST)

**New instance** (Terminal C) or same Terminal A:
```bash
# Before launching, check findings from plan2
cat _FINDINGS.md | grep -A 5 "Plan2 Findings"

# Then launch plan3
# Prompt: "Implement plan3.md: Add share link helpers to main.js..."
# Model: Sonnet
# Expected: ~40 min, ~22k tokens
```

**Critical checkpoint**: At 50% (after ~20 min), implementer checks tokens:
```bash
# They should have ~22k+ tokens remaining
# If yes → continue and complete
# If no → pause, commit, report
```

---

## Stage 3 — Plan 4 (Sequential)

**Start after**: Stage 2 complete (~11:30pm PST)

**New instance** (Terminal D):
```bash
# Before launching, check findings from plan3
cat _FINDINGS.md | grep -A 5 "Plan3 Findings"

# Then launch plan4
# Prompt: "Implement plan4.md: Add 3 API endpoints to main.js..."
# Model: Sonnet
# Expected: ~40 min, ~22k tokens

# CRITICAL: plan3 and plan4 should update _FINDINGS.md
# with complexity assessments and model recommendations for plan5
```

---

## Stage 4 — Plan 5 & 6 (Parallel)

**Start after**: Stage 3 complete (~12:10am PST)

**Before launching**, read findings:
```bash
cat _FINDINGS.md | grep "Recommendation for plan5"
# This determines if plan5 uses Sonnet or Opus
```

**New instance** (Terminal E) — Plan 5:
```bash
# Check model recommendation from plan4/plan3 findings
# If Opus → use --model opus
# If Sonnet → use --model sonnet

# Prompt: "Implement plan5.md: Add QR overlay window to main.js..."
# Expected: ~40 min, ~28k tokens (Sonnet) or ~20k (Opus)
```

**New instance** (Terminal F) — Plan 6:
```bash
# Prompt: "Implement plan6.md: Add Companion module actions..."
# Model: Haiku
# Expected: ~20 min, ~8k tokens
```

---

## Final Checklist

After all 6 plans complete:

- [ ] All plans committed with token counts
- [ ] `_TOKEN_LOG.md` fully populated
- [ ] `_FINDINGS.md` fully populated
- [ ] Total tokens used: ~97k (well under 150k safety margin)
- [ ] All code merged to main branch
- [ ] Ready for Stage 5 (manual integration verification)

```bash
# Final verification
git log --oneline | head -10
# Should show all 6 plan commits

git status
# Should show "nothing to commit, working tree clean"

cat _TOKEN_LOG.md
# Should show all stages completed with token counts
```

---

## Time Estimates

| Stage | Duration | Start Time | Complete Time |
|-------|----------|------------|---------------|
| Stage 0 | 10 min | [Before launch] | [Already done] |
| Stage 1 | 50 min | 10:00pm PST | 10:50pm PST |
| Stage 2 | 40 min | 10:50pm PST | 11:30pm PST |
| Stage 3 | 40 min | 11:30pm PST | 12:10am PST |
| Stage 4 | 60 min | 12:10am PST | 1:10am PST |
| Stage 5 | 30 min | 1:10am PST | 1:40am PST |
| **Total** | **3.5 hrs** | 10:00pm | 1:40am PST |

---

## Important Reminders

📌 **Token tracking is mandatory**
- Each implementer checks tokens at 50% completion
- Each implementer reports tokens in commit message
- If insufficient tokens → pause and use fresh instance

📌 **Findings must be updated** (plans 2, 3, 4)
- Documents code complexity
- Informs model selection for dependent plans
- Especially important: plan5 model choice depends on plan3 & 4 findings

📌 **Each plan can run independently**
- If one runs out of tokens → no blocking
- Next plan starts with fresh 200k tokens
- Prior commits document context

📌 **Commit early and often**
- Don't lose work to token limits
- Clear commit messages enable recovery

---

## Success =

✅ All 6 plans implemented
✅ ~97k tokens used (under 150k safety margin)
✅ All findings and token logs updated
✅ Code ready for Stage 5 (integration verification)

**You've got this! 🚀**

---

## Quick Links

- **Master plan**: `plan0.md`
- **Token checklist**: `TOKEN_TRACKING_CHECKLIST.md`
- **Quick reference**: `PLAN_INDEX.md`
- **Summary**: `ORCHESTRATION_SUMMARY.md`
