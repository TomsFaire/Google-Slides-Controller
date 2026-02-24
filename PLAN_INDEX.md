# PLAN INDEX — Share Link + QR Overlay

Original detailed plan: `plan.md`

## Parallel Execution Strategy

```
PARALLEL #1          PARALLEL #2          PARALLEL #3
├─ plan1.md          ├─ plan2.md           └─ (blocked until plan2/3 done)
│  (PHP site)        │  (Settings UI)
│                    │
│                    └─ plan3.md
│                       (Electron helpers)
│                       [depends on: plan2]
│
│                    └─ plan4.md
│                       (API endpoints)
│                       [depends on: plan3]
│
│                    ├─ plan5.md
│                    │  (QR overlay)
│                    │  [depends on: plan4]
│                    │
│                    └─ plan6.md
│                       (Companion actions)
│                       [depends on: plan4]
```

## Launch Instructions

### Round 1 — Run in Parallel
**Can run simultaneously in separate Claude instances:**

1. **plan1.md** — PHP redirect site (BlueHost)
   - Self-contained, no app dependencies
   - Deploy and test independently

2. **plan2.md** — Desktop settings UI
   - Self-contained renderer/HTML changes
   - No backend logic yet

### Round 2 — Sequential (depends on Round 1)
**After plan2 is complete:**

3. **plan3.md** — Electron main helpers (getShareLink, etc.)
   - Depends on: plan2 (uses prefs structure)
   - Adds helper functions to main.js

### Round 3 — Sequential (depends on Round 2)
**After plan3 is complete:**

4. **plan4.md** — API endpoints (3 new routes)
   - Depends on: plan3 (calls getShareLink())
   - Adds HTTP endpoints to main.js

### Round 4 — Run in Parallel
**After plan4 is complete:**

5. **plan5.md** — QR overlay window (main.js)
   - Depends on: plan4 (API endpoints call showQrOverlay)
   - Adds window + QR generation to main.js

6. **plan6.md** — Companion module actions
   - Depends on: plan4 (calls /api/show-share-qr, /api/hide-share-qr)
   - Modifies companion-module-gslide-opener/actions.js

## Token Budget Per Plan

| Plan | Module | Estimated Lines | Complexity |
|------|--------|-----------------|-----------|
| plan1 | PHP (external) | ~100 | Low |
| plan2 | renderer.js + HTML | ~50–100 | Low |
| plan3 | main.js | ~100–150 | Medium |
| plan4 | main.js | ~100–150 | Medium |
| plan5 | main.js | ~150–200 | Medium |
| plan6 | actions.js | ~80–120 | Low |

Each plan should fit comfortably in a single Claude instance with typical token budgets.

## Verification Checklist

- [ ] plan1 deployed and redirect tested
- [ ] plan2 UI shows all 4 input fields and persists on restart
- [ ] plan3 genShareCode() + getShareLink() work (test with mock)
- [ ] plan4 API endpoints return correct JSON (test with curl)
- [ ] plan5 QR window appears and auto-hides
- [ ] plan6 Companion actions trigger API endpoints

## Files Modified Summary

| File | Plan(s) |
|------|---------|
| main.js | plan3, plan4, plan5 |
| renderer.js | plan2 |
| HTML (settings section) | plan2 |
| package.json | plan5 (add qrcode) |
| companion-module-gslide-opener/actions.js | plan6 |
| (BlueHost PHP site) | plan1 |

## Notes

- Each plan is independent at its scope; verify the dependency chain before launching
- Run plans in the suggested order to minimize backtracking
- Commit between plans for safety (git commits as boundaries)
- If a dependent plan is blocked, check the prior plan's acceptance tests
