# Token Tracking Checklist for Plan Implementers

Quick reference for token management while implementing plans 1–6.

## Before You Start

- [ ] Note your starting token count (Claude session shows this)
- [ ] Verify plan file has token budget section (it does)
- [ ] Read `_FINDINGS.md` to check model recommendations
- [ ] Create/update `_TOKEN_LOG.md` entry for your instance

**Starting tokens**: _____ (out of 200k limit)

---

## During Implementation

### Checkpoint #1: At 25% Completion
- [ ] Check tokens remaining
- [ ] If remaining > (budget × 2) → Continue normally
- [ ] If remaining < budget → Still fine, keep going
- [ ] If remaining < (budget × 0.5) → Monitor closely

### Checkpoint #2: At 50% Completion ⚠️ CRITICAL
- [ ] Check tokens remaining in Claude session
- [ ] Compare to plan budget:
  - plan1: Need 3k remaining? ✓ (cheap)
  - plan2: Need 8k remaining? ✓ (cheap)
  - plan3: Need 18k remaining? **Check now**
  - plan4: Need 18k remaining? **Check now**
  - plan5: Need 22k remaining? **Check now**
  - plan6: Need 5k remaining? ✓ (cheap)

**If tokens too low**:
- [ ] Commit what you have immediately
- [ ] Add message: "plan[N]: 50% complete, insufficient tokens for full implementation"
- [ ] Next instance will start fresh and continue from your notes

### Checkpoint #3: At 75% Completion
- [ ] Check tokens again
- [ ] By this point, you should be mostly done
- [ ] If tokens low, just finish the plan

### Checkpoint #4: Before Committing Final Work
- [ ] Verify tokens used (Claude session shows)
- [ ] Add to commit message:
  ```
  plan[N]: Complete — [actual_tokens]k used, [remaining]k remaining
  ```

---

## Committing Your Work

### Required Commit Message Format

```
plan[N]: Implementation complete

Model used: [Haiku/Sonnet/Opus]
Tokens used: ~[X]k tokens
Tokens remaining: ~[Y]k tokens

Status: [COMPLETED/PAUSED]

[Plan-specific notes if needed]

Findings updated: [YES/NO — only for plans 2-4]
```

### Example

```
plan3: Electron helpers implemented

Model used: Sonnet
Tokens used: ~20k tokens
Tokens remaining: ~20k tokens

Status: COMPLETED

All 3 helpers (genShareCode, registerShareCode, getShareLink)
working correctly with caching logic verified.

Findings updated: YES (_FINDINGS.md includes main.js complexity assessment)
```

---

## After Completion

### Update Central Tracking

1. **Update `_TOKEN_LOG.md`**:
   ```bash
   # Add your section:
   git add _TOKEN_LOG.md
   git commit --amend  # or new commit if appropriate
   ```

2. **If updating `_FINDINGS.md`** (plans 2, 3, 4 only):
   ```bash
   git add _FINDINGS.md
   git commit  # or include in same commit as code
   ```

3. **Verify next stage can start**:
   - [ ] All files committed
   - [ ] `_TOKEN_LOG.md` updated
   - [ ] `_FINDINGS.md` updated (if applicable)
   - [ ] Next plan implementer has clear handoff

---

## Emergency: Low on Tokens

### If you hit < 30k tokens remaining mid-plan:

1. **Finish fast** ⚡
   - Skip extra testing
   - Skip code cleanup
   - Just get it working

2. **Commit immediately** 🔴
   - Add detailed notes
   - Include token count
   - Don't lose work

3. **Next plan starts fresh** 🆕
   - New Claude instance = 200k tokens
   - They'll read your commit for context
   - No loss of progress

### If plan is blocked by token shortage:

1. **Commit exploration/notes**
2. **Update `_TOKEN_LOG.md`**: Mark as "PAUSED"
3. **Add git commit message**: `plan[N]: Incomplete due to insufficient tokens. [Details for resumption]`
4. **Next instance will:**
   - Start with fresh tokens
   - Read your commit
   - Resume from where you left off

---

## Token Budget Quick Reference

| Plan | Model | Budget | Reading | Coding | Testing |
|------|-------|--------|---------|--------|---------|
| plan1 | Haiku | 5k | 1k | 2k | 1k |
| plan2 | Haiku | 10k | 4k | 3k | 2k |
| plan3 | Sonnet | 22k | 10k | 7k | 3k |
| plan4 | Sonnet | 22k | 10k | 7k | 3k |
| plan5 | Sonnet/Opus | 28k (S) / 20k (O) | 12k | 10k | 4k |
| plan6 | Haiku | 8k | 2k | 4k | 1k |

---

## Token Optimization Tips

If running low:
- **Skip reading** full files; request specific line ranges
- **Use Haiku** instead of Sonnet where possible
- **Batch requests** (read multiple files in one request)
- **Skip optional** testing (rely on linting instead)
- **Use _FINDINGS.md** heavily to avoid re-reading main.js

---

## Troubleshooting

| Problem | Cause | Solution |
|---------|-------|----------|
| Tokens running out faster than expected | Large code files (main.js = 7700 lines) | Start new instance, read commits |
| Not enough tokens to complete plan | Underestimated budget or model too weak | Switch to more capable model (Sonnet→Opus) |
| Unsure if tokens sufficient | Checkpoint too late | Check at 50%, don't wait until 75% |
| Next plan can't start (insufficient tokens) | Prior plan used all tokens | All plans can run in separate instances |

---

## Success Metrics

✓ Plan completed without crashes
✓ Tokens properly tracked in `_TOKEN_LOG.md`
✓ `_FINDINGS.md` updated (if applicable)
✓ All code committed with clear messages
✓ Handoff notes clear for dependent plans

Good luck! 🚀
