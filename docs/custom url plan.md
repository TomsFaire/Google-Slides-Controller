# F4: Named Cloudflare Tunnel Support

**Scope:** Code changes only. DNS/domain setup is a separate user task done before using this feature.  
**Assumption:** User has already created a Cloudflare tunnel and has their credentials file and tunnel name ready. The app just needs to launch it correctly and display the right URL.

---

## New Settings Fields

`preferences.json` gets four new fields:

```json
"tunnelMode": "quick",        // "quick" | "named"
"cfTunnelName": "",           // e.g. "slides-controller"
"cfTunnelHostname": "",       // e.g. "slides.reynoldsproduction.com"
"cfCredentialsPath": ""       // path to the .json credentials file
```

The API token/secret never touches `preferences.json` — it stays on disk only as the credentials file `cloudflared` itself uses, which the user points us to.

---

## main.js Changes

This is where most of the logic lives. Current Quick Tunnel launch (`startCloudflaredTunnel()`, ~line 7087):

```js
cloudflaredProcess = spawn(bin, ['tunnel', '--url', origin], { stdio: ['ignore', 'pipe', 'pipe'] })
// then parse stdout for the trycloudflare.com URL
```

### Changes needed

**1. Read `tunnelMode` from preferences when starting the tunnel, then branch the spawn:**

```js
if (tunnelMode === 'named') {
  cloudflaredProcess = spawn(bin, [
    'tunnel', '--credentials-file', cfCredentialsPath, 'run', cfTunnelName
  ], {
    stdio: ['ignore', 'pipe', 'pipe']
  })
} else {
  // existing quick tunnel spawn unchanged (--url, --no-tls-verify if https)
}
```

> **Note:** Use `--credentials-file <path>` as a CLI flag — not a `TUNNEL_CRED_FILE` env variable (cloudflared does not recognize that env var). Also omit `--url` and `--no-tls-verify` for named tunnels; ingress rules come from the Cloudflare-side tunnel config, not the CLI.

**2. Skip stdout URL-parsing for named tunnels.** Currently `main.js` watches `cloudflared` stdout for the random URL and fires an IPC event. For named tunnels the URL is already known — it's `cfTunnelHostname`. After confirming the process started successfully, call `broadcastTunnelUrl(cfTunnelHostname)` immediately instead of waiting for stdout.

**3. Named tunnel ready detection:** Watch stderr/stdout for `"Registered tunnel connection"` (what `cloudflared` prints when the named tunnel connects) rather than the URL pattern used for Quick Tunnels. Fire the ready event then.

**4. No change needed to tunnel stop/restart logic** — `stopCloudflaredTunnel()` uses `proc.kill('SIGTERM')` + SIGKILL fallback, which works the same either way.

**5. New API routes** — add alongside the existing `/api/tunnel-enable` / `/api/tunnel-disable` block (~line 4535) in `main.js` (not a separate webServer.js — that file does not exist; all HTTP routes live in `main.js`):

- `GET /api/tunnel-config` — returns `tunnelMode`, `cfTunnelName`, `cfTunnelHostname`, `cfCredentialsPath`
- `POST /api/tunnel-config` — saves those four fields to preferences; triggers a tunnel restart if the tunnel is currently running

---

## Settings Web UI (HTML/JS in the Web UI served by main.js)

In the WAN Access section, below the existing Quick Tunnel toggle, add:

- **Radio or toggle:** Quick Tunnel / Named Tunnel
- **Named tunnel sub-form** (hidden unless Named is selected):
  - Tunnel name field
  - Public hostname field (the full subdomain URL)
  - Credentials file path field + a **Browse** button (uses Electron's `dialog.showOpenDialog` via IPC)
- **Status indicator:** shows current tunnel URL regardless of mode — Quick Tunnel shows the random URL as today; Named Tunnel shows the configured hostname

---

## preload.js Changes

Follow the existing per-dialog pattern (`showOpenCertDialog`, `showOpenKeyDialog`):

```js
showOpenCredentialsDialog: () => ipcRenderer.invoke('show-open-credentials-dialog')
```

And a matching `ipcMain.handle('show-open-credentials-dialog', ...)` in `main.js` that calls `dialog.showOpenDialog` filtered to `.json` files.

> **Do not** use a generic `select-file` channel — existing handlers each have a named IPC channel, and consistency matters here.

---

## renderer.js Changes

- When tunnel mode is `named`, display `cfTunnelHostname` as the shareable link instead of waiting for the dynamic URL event. The `tunnel-url-changed` IPC event already flows through `broadcastTunnelUrl()` — named mode just calls that with the stored hostname rather than a parsed URL.
- QR code generation should use whichever URL is active — this should work automatically if the QR logic already reads from the displayed URL.

---

## What Doesn't Change

- The bundled `cloudflared` binary — it already supports named tunnel mode
- The PIN authentication system
- The `isWebUiRestrictedTunnelClient` restriction logic — named tunnel traffic also proxies through localhost, so the existing `cloudflaredEnabled` check covers it unchanged
- Quick Tunnel behavior — completely untouched, remains the default

---

## Suggested Implementation Order

1. `preferences.json` schema + defaults
2. `main.js` tunnel launch branching + named tunnel ready detection
3. New API routes in `main.js` (`GET`/`POST /api/tunnel-config`)
4. Settings page UI (form + show/hide logic)
5. `preload.js` + `main.js` file picker IPC (`show-open-credentials-dialog`)
6. `renderer.js` URL display logic
7. Manual test: toggle between modes, verify URLs, verify restart behavior

---

## Validation Notes (from code review)

- `broadcastTunnelUrl` + `tunnel-url-changed` IPC pattern works as-is — just pass `cfTunnelHostname` for named mode
- `get-tunnel-status` IPC (line 3310) returns `{ enabled, url, running }` — `tunnelUrl` will hold `cfTunnelHostname` for named mode, correct with no handler changes
- `stopCloudflaredTunnel()` SIGTERM/SIGKILL pattern works identically for both modes
