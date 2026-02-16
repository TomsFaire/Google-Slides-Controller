# Edits between 1.9.0 and 1.9.1

Speaker-notes text encoding changes have been **rolled back** to the 1.9.0 state. The following summarizes all other changes added since 1.9.0 (for 1.9.1).

---

## 1. Public / tunnel URL

**Files:** `index.html`, `renderer.js`, `main.js` (prefs merge)

- **Pref:** `tunnelPublicUrl` (optional string).
- **Desktop:** In Network Ports / Network Access, added:
  - Input "Public or tunnel URL" (`#tunnel-public-url`).
  - "Share this link: [url]" block (`#share-link-display`, `#share-link-url`) when URL is set.
- **Renderer:** Load/save `tunnelPublicUrl`, `updateShareLinkDisplay()`, `saveTunnelPublicUrl()`.
- **Web UI (main.js):** When `tunnelPublicUrl` is set, header shows "Share this link" bar with URL and Copy button (and small CSS for `.share-link-bar` etc.).

---

## 2. HTTPS for Web UI

**Files:** `main.js`, `index.html`, `renderer.js`, `preload.js`

- **Prefs:** `webUiUseHttps`, `webUiCertPath`, `webUiKeyPath`.
- **main.js:**
  - `https` and `child_process.execSync` required.
  - `getWebUiHttpsCredentials()`: reads user cert/key paths or generates self-signed (OpenSSL) in userData.
  - `startWebUiServer()`: uses `requestHandler`; if credentials exist, `https.createServer(creds, requestHandler)`, else `http.createServer(requestHandler)`; log line uses `protocol` (https/http).
- **Desktop UI:** Checkbox "Serve Web UI over HTTPS", optional cert/key file inputs and Choose/Clear; refs and handlers in renderer; `updateHttpsCertKeyVisibility()`, `saveHttpsPreferences()`.
- **IPC:** `show-open-cert-dialog`, `show-open-key-dialog`; exposed in preload as `showOpenCertDialog`, `showOpenKeyDialog`.

---

## 3. Default port 443 when HTTPS enabled

**Files:** `main.js`, `renderer.js`

- When user enables HTTPS and current Web UI port is 80, port is set to **443** and saved.
- **main.js:** `DEFAULT_WEB_UI_HTTPS_PORT = 443`; if HTTPS and port 80, use 443 before listening.
- **renderer.js:** On HTTPS checkbox change, if port is 80 set to 443; `saveHttpsPreferences()` includes `webUiPort`; `updateNetworkInfo()` uses 443 for Web UI URLs when HTTPS and saved port is 80.

---

## 4. URL display (protocol + tunnel)

**Files:** `renderer.js`

- In `updateNetworkInfo()`, Web UI URLs use `https://` when `preferences.webUiUseHttps` is true.
- Tunnel URL is shown in the "Share this link" block when set (existing behavior).

---

## 5. Documentation

**Files:** `docs/PUBLIC-ACCESS.md`, `README.md`, `README-SECURITY.md`

- **docs/PUBLIC-ACCESS.md:** Public access and HTTPS (tunnels, reverse proxy, in-app HTTPS).
- **README.md:** New section "Public access and HTTPS" with link to that doc.
- **README-SECURITY.md:** Section "If macOS Says the App Is Damaged" and `xattr -cr` for unsigned builds.

---

## 6. Backup (Primary/Backup) UI fix

**Files:** `main.js` (Web UI), `renderer.js`, `index.html`

- **Web UI backup IP row:** Row/input styling so the field is not tiny and typing is visible: `width: 100%`, `minWidth: 0`, `flex: 1 1 0%`, `minWidth: 140px` on input; same for row and `#web-backup-ip-list`.
- **Electron backup IP row:** Same layout (row/input/list) and placeholder "192.168.1.100 or hostname".
- **index.html:** `#backup-ip-list` given `width: 100%; min-width: 0`.

---

## 7. Mac build / quarantine

**Files:** `package.json`, `scripts/after-pack-mac.js`, `README-SECURITY.md`

- **package.json:** Root `build.afterPack` set to `"./scripts/after-pack-mac.js"`.
- **scripts/after-pack-mac.js:** afterPack hook for macOS; runs `xattr -cr` on the built `.app` (only when `context.electronPlatformName === 'darwin'`).
- **README-SECURITY.md:** "If macOS Says the App Is Damaged" and `xattr -cr` instructions (see §5).

---

## Summary of files touched (excluding speaker-notes rollback)

| File | Changes |
|------|--------|
| `main.js` | HTTPS server, credentials, tunnel URL in Web UI HTML, backup IP row styling (Web UI), afterPack not in main (only in package.json) |
| `index.html` | Tunnel URL input, share link block, HTTPS checkbox/cert/key, backup-ip-list width |
| `renderer.js` | Tunnel URL, HTTPS prefs and UI, URL display (https/tunnel), backup IP row styling, saveHttpsPreferences |
| `preload.js` | showOpenCertDialog, showOpenKeyDialog |
| `package.json` | afterPack, build scripts |
| `scripts/after-pack-mac.js` | New file |
| `docs/PUBLIC-ACCESS.md` | New file |
| `docs/EDITS-1.9.0-TO-1.9.1.md` | This file |
| `README.md` | Public access and HTTPS section |
| `README-SECURITY.md` | Damaged-app / xattr section |

Speaker-notes code is back to 1.9.0: server `normalizeSpeakerNotes` (regex only), extraction returns raw text only, Web UI `normalizeSpeakerNotes` (regex, including `\uFFFD+` in source — note: in the embedded template the `\uFFFD` may still be emitted as one character, so client replacement char handling can remain broken at 1.9.0).
