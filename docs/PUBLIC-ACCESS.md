# Remote Access and HTTPS

By default, the Web UI is HTTP-only and available on your local network only. To give remote users access or encrypt the connection, you have a few options.

## Tunnel (easiest for sharing)

Get a public HTTPS URL without needing a domain or port forwarding.

### Built-in WAN access (Cloudflare Quick Tunnel)

The desktop app can start a **Quick Tunnel** from **Settings → WAN Access** (after you run `yarn download:cloudflared` once to fetch bundled binaries).

If the Web UI uses **HTTPS** (including the app’s self-signed certificate), the tunnel runs `cloudflared` with **`--no-tls-verify`** so the origin connection succeeds. Traffic to Cloudflare’s edge is still encrypted.

**Important:** Anyone with the tunnel link can use the web remote until you disable the tunnel or restart the app. Treat the URL like a password—**or** enable an optional **Web UI PIN** in the desktop app (**Settings → WAN Access**) and choose whether it applies to the **tunnel**, **LAN** (non-localhost), or **both**; see the main [README.md](../README.md) WAN section. The **controller IP allowlist** does not limit remote users on this link, because traffic reaches your Web UI from `localhost` via the local `cloudflared` process.

### Restricted Web UI on the shared link

To reduce exposure when you share the Quick Tunnel URL:

- The in-browser **Settings** tab is **not shown** for tunnel users (connections seen as **localhost** from `cloudflared`).
- **Remote** and **Controls** tabs remain available.
- The Web UI’s **API proxy** refuses certain routes in that mode (for example `GET/POST /api/preferences`, `GET /api/displays`, `POST /api/presets`, `POST /api/stagetimer-settings`, and `GET /api/debug/preferences`), so those operations are not available through the shared link alone.

**Administrators:** There is **no** yellow warning banner or similar notice in the Web UI for people using the shared link—they see Remote and Controls only, without being told that Settings exist elsewhere. Rely on this document (and the main [README.md](../README.md) WAN section) when planning access: you need the machine’s **LAN URL** (or equivalent) for full setup.

For **full** Web UI (Settings, saving presets from the browser, etc.), open the app using your machine’s **LAN URL**, not the `trycloudflare.com` link. Note: with Quick Tunnel enabled, opening the Web UI at **`127.0.0.1`** on the same computer also gets the restricted UI; use the LAN IP for local admin.

Direct calls to the **HTTP API** on port **9595** are unchanged and still follow your **controller IP allowlist** (see project security docs).

### Manual Cloudflare Tunnel (CLI)

1. Install [cloudflared](https://developers.cloudflare.com/cloudflare-one/connections/connect-apps/install-and-setup/installation/)
2. Run:
   ```bash
   cloudflared tunnel --url http://localhost:80
   ```
3. Share the generated URL (like `https://xxx.trycloudflare.com`) with remote users

**With ngrok**

1. Install [ngrok](https://ngrok.com/download)
2. Run:
   ```bash
   ngrok http 80
   ```
   (Replace 80 with your custom Web UI port if needed)
3. Share the generated URL (like `https://xxx.ngrok.io`)

## Reverse proxy (for your own domain)

If you have a domain and can run nginx or Caddy on a public server:

1. Point your domain to your server’s IP
2. Configure nginx/Caddy to proxy traffic to `http://YOUR_PRESENTATION_PC_IP:80`
3. Add TLS with Let’s Encrypt or another CA
4. Users access your app at `https://your-domain.com`

No changes needed in the app itself. Check your proxy’s documentation for TLS setup.

## In-app HTTPS (local network only)

Want HTTPS on your LAN? Enable it in Settings → Network Ports → Serve Web UI over HTTPS.

You can:
- Provide a custom certificate and private key (PEM format) from your internal CA
- Leave both empty and let the app generate a self-signed certificate

Browsers will warn you about self-signed certs, but that’s fine for local use. For public access, a tunnel or reverse proxy with a real certificate is better.
