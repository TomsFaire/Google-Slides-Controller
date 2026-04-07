# Remote Access and HTTPS

By default, the Web UI is HTTP-only and available on your local network only. To give remote users access or encrypt the connection, you have a few options.

## Tunnel (easiest for sharing)

Get a public HTTPS URL without needing a domain or port forwarding.

### Built-in WAN access (Cloudflare Quick Tunnel)

The desktop app can start a **Quick Tunnel** from **Settings → WAN Access** (after you run `yarn download:cloudflared` once to fetch bundled binaries).

**Important:** Anyone with the tunnel link can use the web remote until you disable the tunnel or restart the app. Treat the URL like a password. The **controller IP allowlist** does not limit remote users on this link, because traffic reaches your Web UI from `localhost` via the local `cloudflared` process.

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
