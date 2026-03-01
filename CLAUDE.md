# CLAUDE.md

Project: Google Slides Controller (Electron Desktop App)

## Global Development Standards

This project uses the **Everything Claude Code (ECC)** global library for development standards, rules, and workflows:

- **Location:** `~/.claude/skills/` and `~/.claude/rules/`
- **Language-Specific Rules:** `~/.claude/rules/typescript/` (coding style, testing, security, patterns)
- **Common Rules:** `~/.claude/rules/common/` (development workflow, git, testing requirements, security)
- **Available Skills:**
  - `/tdd-workflow` - Test-driven development (80%+ coverage)
  - `/security-review` - Authentication, API security, sensitive data handling
  - `/backend-patterns` - API design, HTTP server patterns, WebSocket handling
  - `/coding-standards` - Code quality and best practices
  - `/verification-loop` - Testing and validation

## Project Overview

**Google Slides Controller** (v1.9.2) is an Electron desktop application for controlling Google Slides presentations across multiple monitors with a web-based remote and HTTP API for AV integration.

- **Stack:** Electron 28, vanilla JavaScript (no production npm deps), electron-builder for packaging
- **Primary Deployment:** macOS (arm64/x64), Linux (AppImage, .deb)
- **API Port:** 9595 (local HTTP/HTTPS server)
- **Key Features:**
  - 4-line SonoBus-style control (future integration planned)
  - Speaker notes capture and display
  - Slide thumbnails and previews
  - Preset management (export/import)
  - Backup/failover for redundancy
  - QR code share link overlay
  - Stagetimer.io integration

## Architecture

Three Electron processes manage the application:

1. **Main Process** (`main.js` ~7700 lines)
   - Window management and lifecycle
   - HTTP/HTTPS API server (port 9595)
   - IPC message handling
   - Google authentication and session persistence
   - Crash reporting and backup/failover logic
   - Stagetimer.io integration

2. **Renderer Process** (`renderer.js` ~1300 lines)
   - Settings UI with DOM bindings
   - Display configuration, presets, debug console

3. **Preload Process** (`preload.js`)
   - IPC security bridge exposing `window.electronAPI` to renderer

## Development Workflow

### Setup

```bash
# Load nvm (required for Node.js)
export NVM_DIR="$HOME/.nvm" && source "$NVM_DIR/nvm.sh"

cd ~/dev/Google-Slides-Controller

# Install dependencies and start
yarn install
yarn start          # Run in development mode
yarn dev            # Watch mode for hot reload
```

### Building

```bash
# Build for specific platform
yarn build:linux    # Build Linux AppImage → dist/
yarn build:win      # Build Windows portable → dist/
yarn build:mac      # Build macOS zip → dist/

# Multi-platform builds
yarn build          # Build all platforms
```

## Key Engineering Patterns

- **Persistent Google Session:** Uses `persist:google` Electron session partition; speaker notes captured via `setupGoogleSessionEncoding()` at main.js:731
- **Speaker Notes:** Scraped from DOM and normalized via `normalizeSpeakerNotes()` at main.js:94 (fixes encoding corruption U+FFFD)
- **Primary/Backup Failover:** Primary broadcasts to backup IPs via `sendToBackups()` at main.js:984; polling via `checkBackupStatus()` at main.js:1035
- **Preferences:** JSON-based config via `loadPreferences()` / `savePreferences()` at main.js:762/794; stored in Electron `userData` path
- **IP Allowlist:** All mutating API calls gated via `isControllerAllowedRequest()` at main.js:950

## Code Standards

Follow TypeScript rules from `/coding-standards`:
- Use immutable data patterns (spread operator, never mutate)
- Keep functions < 50 lines, files < 800 lines
- Comprehensive error handling at all levels
- No hardcoded values (use constants)
- Avoid deep nesting (>4 levels)

## Testing Requirements

- **Minimum 80%+ coverage** (unit + integration tests)
- **Unit tests** for utility functions, API handlers, config serialization
- **Integration tests** for IPC, window lifecycle, Google session
- **Manual acceptance** checklist for Stagetimer.io, failover, WAN scenarios

## API Endpoints

Key HTTP endpoints on local API:

| Method | Endpoint | Purpose |
|--------|----------|---------|
| GET | `/api/status` | Health + current state |
| GET | `/api/get-speaker-notes` | Current slide speaker notes (normalized) |
| GET | `/api/get-slide-previews` | Slide thumbnail images |
| GET | `/api/presets` | Saved presentation presets |
| POST | `/api/open-presentation` | Open presentation URL |
| POST | `/api/next-slide` / `/api/previous-slide` | Navigate slides |
| POST | `/api/go-to-slide` | Jump to specific slide |
| POST | `/api/reload-presentation` | Reload current presentation |
| POST | `/api/share-link` | Generate/get share link with QR |
| POST | `/api/show-share-qr` | Display QR code on presentation screen |

All endpoints require IP allowlist validation; see `isControllerAllowedRequest()` for security details.

## Security Checklist

Before every commit:
- ✅ No hardcoded secrets (API keys, tokens, passwords)
- ✅ All user inputs validated
- ✅ IPC messages properly gated (`preload.js` security bridge)
- ✅ Error messages don't leak sensitive data
- ✅ Session data encrypted at rest (Electron partition)
- ✅ IP allowlist enforced on all mutating endpoints
- ✅ HTTPS support with proper certificate handling

Use `/security-review` skill for authentication and API security questions.

## Deployment Notes

- **macOS:** Both arm64 and x64 builds included in release
- **Linux:** AppImage for universal Linux; .deb packages for Debian-based distros
- **WAN Access:** Requires Cloudflare WARP/VPN overlay (no direct WAN in MVP)
- **Config Backup:** Preferences auto-backup on save; manual export/import supported

## References

- `main.js` – Main process implementation
- `CHANGELOG.md` – Version history and feature changes
- `_FINDINGS.md` – Technical findings and design decisions
- `.claude/settings.local.json` – Local development settings
