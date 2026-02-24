# plan1.md — Tiny PHP Redirect Site (BlueHost)

**Scope**: Part A only
**Dependencies**: None
**Can run in parallel**: Yes

## Goal
Deploy a minimal PHP redirect service on BlueHost that maps short codes to long URLs with TTL.

## Endpoints

### POST /api/register
```
Body: { code, target, ttlSec, key }
- Auth: key must match server-side secret
- Store: code -> target + expiresAt
- Return: { ok:true, url:"https://yourdomain/<code>", expiresAt }
```

### GET /<code>
```
- Look up mapping; if missing/expired => 404
- If found => 302 redirect to target
- Auto-cleanup expired rows on lookup
```

## Implementation Notes
- Storage: SQLite file (preferred) or flat JSON
- Cleanup: delete expired rows on register + on redirect lookup
- Keep it minimal (~100 lines of PHP)

## Acceptance Checks
- Register a code with ttl=300
- Verify redirect works
- Verify redirect returns 404 after TTL expires
- Verify cleanup removes expired rows

---

## Token Tracking

**Estimated budget**: 5k tokens (3–8k range)

**Before starting**: Note available tokens in your Claude session.

**During implementation**:
- Design simple PHP structure — ~1–2k tokens
- Write PHP code for /api/register endpoint — ~2–3k tokens
- Write PHP code for redirect logic — ~1–2k tokens
- Plan deployment + testing — ~1k tokens

**This is the simplest plan**, so token tracking is minimal.

**Before completing**:
1. Check tokens remaining
2. If tokens remaining < 8k:
   - **Finish plan1** (it's small, nearly done)
   - Commit with token count
3. If tokens remaining > 15k:
   - Safe to continue with parallel plan2

**After completion**:
Add this to commit message and/or `_TOKEN_LOG.md`:
```
plan1: ~4k tokens used, X tokens remaining
PHP redirect site: [status — deployed/ready to deploy]
Status: READY FOR stage 2
```

**Note**: plan1 runs in parallel with plan2, so token usage doesn't block anything.
Both should complete within 50 minutes combined.
