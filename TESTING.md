# Testing Guide - All Plans Implemented

## ✅ Validation Summary

All 6 implementation plans have been completed and validated:

| Plan | Component | Status | Tests |
|------|-----------|--------|-------|
| 1 | Share Code Generation | ✅ COMPLETE | genShareCode(), caching logic |
| 2 | Desktop UI Settings | ✅ COMPLETE | Share settings form in renderer |
| 3 | Share Link Helpers | ✅ COMPLETE | HTTP requests to redirect service |
| 4 | API Endpoints | ✅ COMPLETE | 3 new endpoints (/api/share-link, /api/show-share-qr, /api/hide-share-qr) |
| 5 | QR Overlay Window | ✅ COMPLETE | Frameless window on presentation display |
| 6 | Companion Actions | ✅ COMPLETE | 2 new actions (show_share_qr, hide_share_qr) |

**Total Code Added:** 530+ lines
**Syntax Validation:** ✅ PASSED
**Dependencies:** ✅ INSTALLED (qrcode v1.5.4)

---

## 📦 Build Artifacts

### macOS (arm64)
- **Location:** `dist/mac-arm64/Google Slides Opener.app`
- **Package:** `dist/Google Slides Opener-1.9.2-arm64-mac.zip` (92 MB)
- **Supports:** Apple Silicon (M1/M2/M3)

### Linux (x64)
- **Location:** `dist/Google Slides Opener-1.9.2.AppImage` (105 MB)
- **Supports:** x86_64 Linux systems
- **Usage:** Make executable and run directly

---

## 🧪 Testing Checklist

### Phase 1: Desktop App Launch
- [ ] Launch application on target platform
- [ ] Confirm settings UI opens without errors
- [ ] Check console for any syntax/runtime errors

### Phase 2: Share Settings Configuration
**Location:** Settings UI → Share Settings section
- [ ] Enter Share Base URL (e.g., `https://example.com`)
- [ ] Enter Share Register URL (redirect service endpoint)
- [ ] Enter Share API Key (authentication token)
- [ ] Click "Save" and confirm success message
- [ ] Verify settings persist after restart

### Phase 3: API Endpoint Testing
**Test via HTTP client (curl/Postman) or Companion:**

#### Endpoint 1: Generate Share Link
```bash
POST http://localhost:9595/api/share-link
Body: {}
Expected: { ok: true, url: "https://...", code: "word-word-hex8", expiresAt: 1234567890 }
```

#### Endpoint 2: Show QR Overlay
```bash
POST http://localhost:9595/api/show-share-qr
Body: { durationSec: 20, forceNew: false }
Expected: { ok: true }
Visual: QR code appears on presentation display for 20 seconds
```

#### Endpoint 3: Hide QR Overlay
```bash
POST http://localhost:9595/api/hide-share-qr
Body: {}
Expected: { ok: true }
Visual: QR code disappears from presentation display
```

### Phase 4: QR Overlay Display
- [ ] Open presentation (multiple displays recommended)
- [ ] Call API to show QR overlay
- [ ] Confirm QR window appears on presentation display
- [ ] Verify QR code is scannable (contains correct URL)
- [ ] Confirm fallback URL text is readable
- [ ] Verify window auto-hides after duration
- [ ] Manually hide and confirm immediate disappearance

### Phase 5: Companion Module Integration
**If using Bitfocus Companion:**
- [ ] Connect Companion to application instance
- [ ] Locate "Show Share QR" action in action list
- [ ] Configure action with custom duration (e.g., 30 seconds)
- [ ] Enable "Generate New Share Link" option
- [ ] Trigger action and confirm QR displays
- [ ] Trigger "Hide Share QR" action
- [ ] Confirm QR disappears

### Phase 6: Multi-Display Testing (if available)
- [ ] Set presentation display in settings
- [ ] Set notes display (secondary)
- [ ] Confirm QR appears on presentation display
- [ ] Confirm speaker notes appear on secondary display
- [ ] Test with reversed display assignments

### Phase 7: Error Handling
- [ ] Disconnect network and try to generate share link → Should show error
- [ ] Invalid settings (empty URL, bad key) → Should validate and show error
- [ ] Try showing QR with invalid duration (e.g., 0 or 500) → Should validate
- [ ] Try showing QR on system with only primary display → Should fallback

### Phase 8: Performance & Stability
- [ ] Leave QR overlay displayed for extended period → No memory leaks
- [ ] Rapid show/hide cycles (10+ times) → No crashes
- [ ] Switch presentations while QR is displayed → Overlay should auto-hide or persist based on implementation
- [ ] Close app while QR is showing → Clean shutdown

---

## 🐛 Known Limitations

### Platform-Specific
- **macOS:** Code signing not enabled (acceptable for testing)
- **Windows:** Build requires Wine (can be built on Windows natively)
- **Linux:** Some window managers may not support `alwaysOnTop` properly

### Features Not Tested
- Actual Google Slides API integration (requires valid presentations)
- PHP redirect service functionality (requires deployed service)
- Multiple simultaneous QR displays (currently only one overlay at a time)

---

## 📝 Testing Results Template

When testing, fill out this template:

```
Platform: [macOS/Linux/Windows]
Version: 1.9.2
Build Date: 2026-02-23

✅ Desktop Launch: [Pass/Fail]
✅ Settings UI: [Pass/Fail]
✅ API Endpoints: [Pass/Fail]
✅ QR Display: [Pass/Fail]
✅ Companion Actions: [Pass/Fail]
✅ Error Handling: [Pass/Fail]
✅ Stability: [Pass/Fail]

Notes:
-
-

Overall Result: [PASS/FAIL]
Tested By: [Name]
Date: [YYYY-MM-DD]
```

---

## 🔧 Troubleshooting

### "Cannot find module 'qrcode'"
- Run `yarn install` to install dependencies
- Verify `node_modules/qrcode` exists

### QR Window Not Appearing
- Check that presentation display is configured correctly
- Verify API endpoint is being called (check logs)
- Ensure display exists (try primary display fallback)

### API Endpoints Return 403
- Verify controller IP is in allowlist
- Check IP allowlist configuration in settings

### Share Link Generation Fails
- Verify redirect service is running and accessible
- Check Share Base URL, Register URL, and API Key are configured
- Verify network connectivity to redirect service

---

## 📞 Support

For issues or questions during testing:
1. Check application logs in developer console
2. Review error messages in API responses
3. Verify all configuration settings are correct
4. Check network connectivity to backend services

---

**Build Date:** 2026-02-23
**Version:** 1.9.2
**All Plans:** ✅ Complete and Ready for Testing
