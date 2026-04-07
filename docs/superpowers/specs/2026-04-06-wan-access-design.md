# WAN Access via Cloudflare Quick Tunnels

## Goal

Ship an **A+** implementation: reproducible builds, correct packaging, safe operator expectations, Windows parity, and clean process lifecycle.

---

## Context

Remote viewers and operators need the Google Slides web remote **without** being on the same LAN as the presentation machine. The Web UI server already proxies `/api/*` to the main API (search `startWebUiServer` and the `/api` proxy block in `main.js`), so one tunnel aimed at the **Web UI port** is enough—no second tunnel.

**Approach:** Bundle `cloudflared` binaries, spawn a child process (`tunnel --url http://localhost:<webUiPort>`), parse the `trycloudflare.com` URL from stdout/stderr, show it in Settings with copy.

**Limitations (Option A, accepted):** Random URL per run; share before each event; no Cloudflare account or API key.

**Line numbers elsewhere in this doc are hints only**—search by symbol name if they drift.

---

## Security & threat model (read before coding)

The Web UI handler applies `isControllerAllowedRequest()` to **every** request (static pages and `/api` proxy). That check uses `req.socket.remoteAddress`.

Traffic from a Quick Tunnel flows: **Internet → Cloudflare → `cloudflared` on the host → `http://127.0.0.1:<webUiPort>`**. From Node’s perspective the connection is **from localhost**, and **`isLocalhostAddress()` always allows those requests**.

**Implication:** The controller IP allowlist **does not restrict** who can use the app **through an active tunnel**. Anyone who obtains the tunnel URL gets the same access as a local browser session to the Web UI (including slide control via the UI).

**Required product behavior:**

1. **Settings copy:** Clear warning that the tunnel URL is **secret** (password-equivalent); anyone with the link can operate the remote for the life of the tunnel.
2. **Optional checkbox:** “I understand the tunnel URL must not be posted publicly” before first enable (stored in prefs once accepted), or a single persistent hint in the WAN card—pick one; don’t ship without visible warning.
3. **Docs:** Add a short paragraph to `README.md` or `docs/PUBLIC-ACCESS.md` cross-linking this behavior.

**Future hardening (out of scope for v1 unless time allows):** tunnel-only shared secret (path prefix or header), or a dedicated “read-only tunnel” mode that blocks proxied mutating `/api` routes.

---

## Platform matrix

| Platform | cloudflared artifact | Packaged path |
|----------|----------------------|---------------|
| macOS arm64 | `cloudflared-darwin-arm64` | `Resources/cloudflared/…` |
| macOS x64 | `cloudflared-darwin-amd64` | same |
| Linux arm64 | `cloudflared-linux-arm64` | same |
| Linux x64 | `cloudflared-linux-amd64` | same |
| Windows x64 | `cloudflared-windows-amd64.exe` | same folder under `resources` |

**Windows:** Include the `.exe` in download script and `getCloudflaredBinaryPath()`. Use `spawn` with the `.exe` name; no `chmod` on Windows.

---

## Files to Modify

| File | Change |
|------|--------|
| `main.js` | Tunnel process management, binary resolution, guards, IPC, lifecycle, robust stop |
| `preload.js` | Expose tunnel IPC + event (same security model as existing prefs IPC) |
| `index.html` | WAN Access section + security hint / acknowledgement |
| `renderer.js` | Load/save, wiring, copy, `onTunnelUrlChanged` |
| `package.json` | **Merge** `extraResources` (keep existing entries) |
| `scripts/download-cloudflared.sh` | *(new)* Pin version; fetch all platform binaries |
| `.gitignore` | `resources/cloudflared/` (already specified in repo) |
| `README.md` or `docs/PUBLIC-ACCESS.md` | One subsection on Quick Tunnel + allowlist caveat |

---

## Step 1 — Download script (pinned version, all platforms)

**Why pin:** `releases/latest` breaks reproducible builds and CI.

Create `scripts/download-cloudflared.sh`:

```bash
#!/bin/bash
set -euo pipefail
# Pin a version; bump intentionally when upgrading cloudflared.
VERSION="${CLOUDFLARED_VERSION:-2026.3.0}"
mkdir -p resources/cloudflared
BASE="https://github.com/cloudflare/cloudflared/releases/download/${VERSION}"
declare -a FILES=(
  cloudflared-darwin-amd64
  cloudflared-darwin-arm64
  cloudflared-linux-amd64
  cloudflared-linux-arm64
  cloudflared-windows-amd64.exe
)
for f in "${FILES[@]}"; do
  echo "Downloading $f ..."
  curl -fL -o "resources/cloudflared/$f" "$BASE/$f"
done
chmod +x resources/cloudflared/cloudflared-darwin-* resources/cloudflared/cloudflared-linux-* 2>/dev/null || true
echo "Done. Version ${VERSION}"
```

- Run: `bash scripts/download-cloudflared.sh`
- Override: `CLOUDFLARED_VERSION=x.y.z bash scripts/download-cloudflared.sh`
- Ensure `resources/cloudflared/` is gitignored (~30 MB per file).

---

## Step 2 — electron-builder: merge `extraResources`

**Do not replace** the existing `extraResources` array. Today the app includes e.g. `BUILD-INFO.txt`. **Append** the cloudflared bundle:

```json
"extraResources": [
  "BUILD-INFO.txt",
  {
    "from": "resources/cloudflared",
    "to": "cloudflared",
    "filter": ["**/*"]
  }
]
```

Packaged layout: `process.resourcesPath/cloudflared/<binary-name>`.

**CI / release builds:** Document that `download-cloudflared.sh` must run **before** `electron-builder` if `resources/cloudflared` is empty (add a step to `.github/workflows/build.yml` or fail the build with a clear error).

---

## Step 3 — `main.js`

### 3a. Globals

Near other server globals:

```javascript
let cloudflaredProcess = null;
let tunnelUrl = null;
let cloudflaredKillTimer = null;
```

### 3b. Binary path helper

- Use `app.isPackaged ? process.resourcesPath : path.join(__dirname, 'resources')`.
- **Darwin:** `cloudflared-darwin-${arch}` with `arch` from `process.arch` (`arm64` vs `amd64` for x64).
- **Linux:** `cloudflared-linux-${arch}`.
- **Windows:** `process.platform === 'win32'` → `cloudflared-windows-amd64.exe`.

### 3c. Preconditions before `spawn`

```javascript
function getCloudflaredBinaryPath() { /* as above */ }

function assertCloudflaredAvailable() {
  const bin = getCloudflaredBinaryPath();
  if (!fs.existsSync(bin)) {
    logError('[Tunnel] Binary missing:', bin, '— run scripts/download-cloudflared.sh');
    return null;
  }
  return bin;
}
```

If missing, log, notify renderer (optional IPC event `tunnel-error`), and do not leave the UI stuck on “Connecting…” without explanation.

### 3d. URL parsing (primary + fallback)

```javascript
const URL_RE_STRICT = /https:\/\/[a-z0-9-]+\.trycloudflare\.com\b/;
const URL_RE_LOOSE = /https:\/\/[^\s"'<>]+\.trycloudflare\.com\b/i;

function extractTunnelUrl(chunk) {
  const text = chunk.toString();
  let m = text.match(URL_RE_STRICT);
  if (m) return m[0];
  m = text.match(URL_RE_LOOSE);
  return m ? m[0] : null;
}
```

Log the first few stderr lines at `debug` if no URL after ~30s (helps field support).

### 3e. Start / stop with reliable teardown

```javascript
function stopCloudflaredTunnel() {
  if (cloudflaredKillTimer) {
    clearTimeout(cloudflaredKillTimer);
    cloudflaredKillTimer = null;
  }
  if (!cloudflaredProcess) return;
  const proc = cloudflaredProcess;
  cloudflaredProcess = null;
  tunnelUrl = null;
  try {
    proc.kill('SIGTERM');
  } catch (e) { /* ignore */ }
  cloudflaredKillTimer = setTimeout(() => {
    cloudflaredKillTimer = null;
    try {
      if (!proc.killed) proc.kill('SIGKILL');
    } catch (e) { /* ignore */ }
  }, 5000);
  broadcastTunnelUrl(null);
}

function broadcastTunnelUrl(url) {
  BrowserWindow.getAllWindows().forEach(w =>
    w.webContents.send('tunnel-url-changed', url)
  );
}
```

On `exit`, clear `cloudflaredKillTimer`, null the process handle, broadcast `null`.

### 3f. `startCloudflaredTunnel`

- If `!loadPreferences().cloudflaredEnabled` return.
- `const bin = assertCloudflaredAvailable(); if (!bin) return;`
- After `startWebUiServer()` has bound the port, spawn (otherwise you may race).
- Attach `stdout`/`stderr` handlers using `extractTunnelUrl`; set `tunnelUrl` once and broadcast.

### 3g. IPC

```javascript
ipcMain.handle('get-tunnel-status', () => ({
  enabled: !!loadPreferences().cloudflaredEnabled,
  url: tunnelUrl,
  running: !!cloudflaredProcess
}));

ipcMain.handle('set-tunnel-enabled', async (_event, enabled) => {
  const prefs = loadPreferences();
  prefs.cloudflaredEnabled = !!enabled;
  savePreferences(prefs);
  if (enabled && !cloudflaredProcess) startCloudflaredTunnel();
  else if (!enabled) stopCloudflaredTunnel();
  return { success: true };
});
```

Use the same IPC exposure pattern as other settings (preload whitelist only).

### 3h. Lifecycle

- `app.whenReady()`: after `startWebUiServer()` succeeds, call `startCloudflaredTunnel()`.
- `app.on('before-quit')`: call `stopCloudflaredTunnel()` (await or sync kill; prefer stopping tunnel before destroying windows if order matters).

### 3i. Preference key

- Add `cloudflaredEnabled` (boolean, default `false`) to preference load/save paths if you centralize defaults—keep consistent with `loadPreferences` / `savePreferences`.

---

## Step 4 — `preload.js`

Expose:

```javascript
getTunnelStatus: () => ipcRenderer.invoke('get-tunnel-status'),
setTunnelEnabled: (enabled) => ipcRenderer.invoke('set-tunnel-enabled', enabled),
onTunnelUrlChanged: (callback) => {
  if (typeof callback !== 'function') return;
  ipcRenderer.on('tunnel-url-changed', (_event, url) => callback(url));
},
```

---

## Step 5 — `index.html`

- WAN section as in the original spec.
- **Add** a visible warning (e.g. `small.field-hint` or alert-style):

  > Anyone with this link can control the web remote until you turn the tunnel off or restart the app. The link is random each session—treat it like a password. Your controller IP allowlist does not apply to traffic through this tunnel.

- Optional: one-time “I understand” checkbox stored in prefs (`wanTunnelRiskAcknowledged`) gating the enable checkbox—implementer’s choice; at minimum static warning text.

---

## Step 6 — `renderer.js`

Same wiring as the original spec (`loadTunnelStatus`, checkbox, copy, `onTunnelUrlChanged`).

**Enhancement:** If `get-tunnel-status` returns `enabled && running && !url` for >30s, show “Still connecting… check logs” or surface main-process log hint—optional polish.

---

## Verification checklist (A+ bar)

1. **Bootstrap:** Run `download-cloudflared.sh` — all five artifacts present; sizes non-zero.
2. **Dev:** `yarn start` — enable WAN; URL appears within ~5–15s; copy works; remote browser (off-LAN) loads Web UI and controls slides.
3. **Disable:** Toggle off — process gone (`ps` / Task Manager); URL cleared.
4. **Quit:** No orphaned `cloudflared` after quit (SIGTERM path; SIGKILL only if needed).
5. **Missing binary:** Temporarily rename `resources/cloudflared` — enable tunnel shows error path, no silent hang.
6. **Package:** `yarn build:mac` / `build:linux` / `build:win` — `Resources` (or equivalent) contains `cloudflared/` with correct binary for that arch.
7. **extraResources:** Packaged app still contains prior resources (e.g. `BUILD-INFO.txt`).
8. **Security doc:** Operator can read warning in UI; README/PUBLIC-ACCESS mentions tunnel URL secrecy and allowlist limitation.
9. **Regression:** Direct LAN access to Web UI still respects controller allowlist when configured (unchanged code paths).

---

## Implementation order (suggested)

1. Script + `.gitignore` + merged `extraResources` + CI pre-build step (or build-time check).
2. `main.js` path helper, exists check, spawn, parse, stop with timeout.
3. IPC + preload.
4. UI + renderer + prefs default.
5. Docs + verification pass on two platforms (e.g. macOS + one of Linux/Windows).

---

## Optional follow-ups (not required for A+)

- Metrics: log tunnel connect latency once.
- Named Cloudflare Tunnel (account) as Option B doc.
- Tunnel-only shared secret path (e.g. `/secret-<uuid>/`) enforced in `requestHandler`.
