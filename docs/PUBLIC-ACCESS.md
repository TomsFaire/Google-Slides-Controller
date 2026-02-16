# Public access and HTTPS

The Web UI is served on your local network (e.g. `http://192.168.x.x:80`). By default it is **HTTP only** and not reachable from the internet: firewalls and NAT block direct access, and HTTP is unencrypted.

## Options for remote or secure access

### 1. Tunnel (recommended for a shareable public link)

A tunnel gives you a public HTTPS URL that forwards to your Web UI. No port forwarding or domain required.

**Cloudflare Tunnel**

1. Install [cloudflared](https://developers.cloudflare.com/cloudflare-one/connections/connect-apps/install-and-setup/installation/).
2. Run a quick tunnel to your Web UI port (default 80):
   ```bash
   cloudflared tunnel --url http://localhost:80
   ```
3. Use the printed URL (e.g. `https://xxx.trycloudflare.com`) and optionally paste it into the app’s **Public or tunnel URL** field in Network Access so it’s easy to share.

**ngrok**

1. Install [ngrok](https://ngrok.com/download).
2. Run:
   ```bash
   ngrok http 80
   ```
   (Use your Web UI port if you changed it.)
3. Use the generated `https://xxx.ngrok.io` URL and optionally paste it into the app’s **Public or tunnel URL** field.

### 2. Reverse proxy (if you have a domain and server)

Run nginx or Caddy on a machine with a public IP or domain. Terminate TLS (e.g. with Let’s Encrypt) and proxy to your presentation machine:

- **Upstream:** `http://PRESENTATION_MACHINE_IP:WEB_UI_PORT`
- **Public URL:** `https://your-domain.com`

Point users to your domain. No change is required inside the app. See your proxy’s documentation for TLS and reverse-proxy setup.

### 3. In-app HTTPS (LAN only)

The app can optionally serve the Web UI over HTTPS on your LAN:

- **Settings → Network Ports → Serve Web UI over HTTPS**: enable, then either:
  - **Custom certificate:** choose PEM certificate and private key files (e.g. from your CA or internal PKI), or
  - **Self-signed:** leave cert/key empty; the app will generate a self-signed certificate in its user data directory and reuse it on restart.

Browsers will show a warning for self-signed certs; you can accept it for LAN use. For **public** or trusted HTTPS, a tunnel or reverse proxy usually provides a trusted certificate (e.g. Let’s Encrypt) and is the preferred approach.
