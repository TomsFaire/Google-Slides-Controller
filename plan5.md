# plan5.md — QR Overlay Window (Presentation Display)

**Scope**: Part E only
**Dependencies**: plan4 (API endpoints must exist; they call showQrOverlay/hideQrOverlay)
**Parallel**: No — must wait for plan4

## Goal
Create a frameless, transparent, always-on-top QR overlay window on the Presentation display.

## Implementation

### Add Global in main.js
```javascript
let qrOverlayWindow = null;
```

### showQrOverlay(shareUrl, durationSec = 20)
```
Logic:
  1. Generate QR image from shareUrl (use 'qrcode' npm package → data URL)
  2. Create frameless, transparent Electron window
     - Display: prefs.presentationDisplayId (fallback to primary)
     - Frameless: true
     - Transparent: true
     - AlwaysOnTop: true
     - VisibleOnAllWorkspaces: true (macOS + Windows)
  3. Load data: URL with QR image + URL text fallback
  4. Auto-hide after durationSec (setTimeout → hideQrOverlay())
  5. Set qrOverlayWindow = newWindow
```

### hideQrOverlay()
```
Logic:
  1. If qrOverlayWindow exists:
     - Clear any pending hide timeout
     - window.close()
     - qrOverlayWindow = null
  2. Return ok
```

### Window Content (data: URL)
HTML/CSS in a data: URL:
```
- Big centered QR image
- Below QR: the share URL as plain text (fallback)
- Dark background, white text
- Keep it minimal (< 200 lines)
```

## Dependencies
- `qrcode` npm package: generate QR data URLs
- Add to package.json if not present

## Implementation Notes
- Use existing display selection logic (presentationDisplayId fallback pattern)
- Window should be ~400x400px (big enough for scanning)
- Auto-close timer must be clearable (if hideQrOverlay called early)
- Test on macOS, Windows, Linux display detection

## Acceptance Checks
- QR window appears on Presentation display on showQrOverlay() call
- QR contains correct URL (test by scanning)
- URL text is readable as fallback
- Window auto-hides after durationSec
- hideQrOverlay() hides immediately
- No crashes on missing display

---

## Token Tracking

**Estimated budget**: 28k tokens (Sonnet) / 20k tokens (Haiku)
**Model choice**: Check `_FINDINGS.md` for recommendation (may be Opus due to complexity)

**Before starting**:
- Note available tokens in your Claude session
- Read `_FINDINGS.md` to confirm model (Sonnet, Opus, or Haiku)

**During implementation**:
- Read main.js for window creation patterns — expect ~10–14k tokens
  - Focus on: existing window creation, display selection, frameless windows
  - Less reading if _FINDINGS.md documented patterns well
- Understand data: URL generation + QR library — ~3–4k tokens
- Implement showQrOverlay() + hideQrOverlay() — ~7–10k tokens
- Set up qrcode npm dependency — ~1–2k tokens
- Test window appears/disappears — ~2–3k tokens

**Critical checkpoint at 50% through**:
After reading main.js window patterns (~12k tokens used):
1. Check tokens remaining
2. If remaining > 28k → Safe to continue
3. If remaining < 14k → **PAUSE and report**
   - Commit exploration code
   - Message: "plan5: Window patterns identified, need fresh instance to complete"
   - Plan will continue in new session

**Before completing**:
1. Check tokens remaining
2. If tokens remaining < 28k but > 15k:
   - **Finish plan5** (you're past halfway)
   - Complete implementation
   - Commit with token count
3. If tokens remaining < 10k:
   - Finish and commit
   - That's OK, plan6 is independent

**After completion**:
Add this to commit message and/or `_TOKEN_LOG.md`:
```
plan5: ~25k tokens used, X tokens remaining
QR overlay functions implemented and tested
Status: READY FOR plan6 (plan6 is independent)
```

**If model was upgraded to Opus**:
- More tokens used, but higher quality + faster
- Higher cost, but might save time on debugging
- Note this in commit: "plan5: Implemented with Opus (due to high complexity noted in _FINDINGS.md)"

**If tokens critically low** (< 10k remaining):
- Commit everything
- plan6 can start in fresh instance (independent of plan5 tokens)
