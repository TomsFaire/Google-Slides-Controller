# WaveShare + PerfectCue — testing guide

This guide covers bench validation of the Google Slides Controller PerfectCue path with a **WaveShare** serial–Ethernet adapter. It assumes IP/network parameters are already correct on the converter.

**Related:** [`waveshare-perfectcue-setup.md`](waveshare-perfectcue-setup.md) (network and field checklist).

---

## Safety and pinouts (read first)

1. **PerfectCue RJ45 ports are model-specific.** Some DSAN ports carry **inter-unit cue data and/or power** (see DSAN PerfectCue / cue-light documentation). **Do not** assume “standard Ethernet” T568A/B pinout or plug an unknown homemade cable into audio gear or a switch.
2. **Confirm the exact jack you use.** Your integration uses the **RS485 / network extender path** as discussed with DSAN—verify in the **manual for your PerfectCue model** which RJ45 pins are **RS485 A/B** (and whether **GND** or **shield** is required). If documentation is unclear, contact DSAN support with your model number before applying DC or long cables.
3. **WaveShare form factor.** Many WaveShare **RS232 RS485 TO POE ETH (B)** units expose **RS485 on screw terminals** (`A` / `B` / `GND`) and **RS232 on a DE-9 (DB9)** connector—not RS485 on DB9. **Identify your unit’s silkscreen** before building a cable. If RS485 is only on terminals, use a **terminal block or ferrule** from your RJ45 pigtail, not a DB9 shell, unless you use an approved adapter.

---

## 1. Building the cable (RJ45 on PerfectCue → WaveShare)

### What you need

- DSAN documentation for **your** RJ45 RS485 pin assignment (A, B, optional GND).
- Multi-conductor cable: preferably **one twisted pair** for RS485 (plus drain/GND if specified), length as short as practical for first tests.
- Appropriate termination for the WaveShare side: **screw terminals** (typical) or **DE-9** only if your hardware doc says RS485 is on that connector.
- Optional: **120 Ω** termination resistor across A–B **only if** the DSAN / WaveShare documentation calls for it at this segment (many short bench links do not need it).

### Wiring steps (procedure)

1. **Write down** from the DSAN doc: which RJ45 pin numbers are **RS485+ / A**, **RS485− / B**, and **reference/GND** (if required).
2. On the WaveShare, confirm labels **A / B / GND** (screw terminals) or the DB9 pin numbers **from the WaveShare manual for your SKU**.
3. Connect **A ↔ A** and **B ↔ B** between PerfectCue and WaveShare (polarity swap will prevent reliable data—swap one pair if you see garbage in logs after baud checks).
4. **GND / shield:** Bond signal ground only as specified by both vendors; avoid creating ground loops between unrelated racks.
5. **First power-on:** With the WaveShare powered and PerfectCue on, use the app **debug console** (see section 4) and press **forward/back** on the remote—look for `[PerfectCue] raw:` lines with plausible hex (not all `ff` or repeating noise).

If you cannot obtain a definitive DSAN pinout, stop and obtain it from DSAN; guessing RJ45 assignments risks equipment damage where **12 V or custom signaling** is present on some products.

---

## 2. WaveShare serial settings (after IP is set)

Configure via **web UI** (`http://<device-ip>/`) and/or **Vircom** so they match **what PerfectCue outputs on that RS485 link** and match the **Google Slides Controller** listener.

### Typical values (confirm against your hardware)

- **Baud rate:** **115200** — common on WaveShare paths; DSAN USR modules often used **9600**—if raw bytes look wrong, try the other.
- **Data bits:** **8**
- **Parity:** **None**
- **Stop bits:** **1**
- **Flow control:** **None**

Apply / submit settings and **restart the device** if the UI requires it.

### App-side listener

- In **Settings → PerfectCue**, set the row’s **Converter** dropdown to **WaveShare** for longer keep-alive presets (vs **DSAN**).
- **Save** PerfectCue settings.
- Ensure the WaveShare is still in **TCP Client** mode to the controller’s IP and the **same TCP port** as that row (default **8899**).

---

## 3. WaveShare keep-alive / idle timers

Goal: the converter must stay connected through idle periods **without** requiring you to poll manually; values should be **consistent** with the app’s adapter preset (see `src/perfectcue-adapter-presets.js` on the **WaveShare** branch).

### Web UI (example labels)

- **Reconnect time:** A modest value (e.g. **12 s**) helps recover after the Electron app restarts without power-cycling hardware.
- **No-data restart:** Prefer **Disabled** for long presentations with no button presses.
- **Keep-alive / TCP idle** (wording varies): Set so the firmware does **not** drop the TCP session **before** the app’s next application ping. For **WaveShare** preset the app sends **`0xFF` about every 45 s** and uses a longer read-idle timeout than DSAN—your device **keep-alive** should be **≥ ~60 s** or aligned with WaveShare’s docs so it does not fight the app (see project plan: avoid firmware closing at 30 s while the app pings every 45 s).

After changes, **submit** and reboot the converter if required.

### Vircom

If you use **Vircom** instead of the browser: open the device → **Advanced Settings** → adjust **Keep Alive Time**, **Reconnect Time**, and disable **Restart for no data** if present → **Modify Setting** / restart as prompted.

---

## 4. Building and running the test app for production machines (macOS)

Distribute a **release build** from the branch that contains the WaveShare work (e.g. **`WaveShare`** merged to `main`, or build directly from **`WaveShare`** for a field trial).

### Prerequisites (build machine)

- **Node + Yarn** (see project [`CLAUDE.md`](../CLAUDE.md) for `nvm` if needed).
- **macOS** for `.dmg` / `.zip` outputs targeting Mac production machines.

### Commands

```bash
cd /path/to/gslide-opener
yarn install
yarn build:mac
```

### Artifacts (electron-builder)

Outputs go under **`dist/`**, for example:

- **`Google Slides Opener-<version>-arm64.dmg`** — Apple Silicon Macs  
- **`Google Slides Opener-<version>-arm64-mac.zip`** — zip variant for deployment  
- **`Google Slides Opener-<version>-x64.dmg`** / **`-mac.zip`** — Intel Macs  

Ship the **architecture that matches** each production Mac (`uname -m`: `arm64` vs `x86_64`). Zip files are convenient for policies that block `.dmg`; DMG is typical for drag-to-Applications installs.

### Runtime test procedure

1. Install or unzip the build on the **controller Mac**.
2. Open **Settings**, enable **PerfectCue**, set listener port(s), set converter type **WaveShare** on the WaveShare line, **Save**.
3. Open **Debug** (or the debug console) and confirm:
   - `[PerfectCue] Listening on port <port>`
   - After the WaveShare connects: `client connected (waveshare) from <ip>`
4. Press **forward / back** on the remote; confirm **`[PerfectCue] raw:`** and slide changes.
5. **Idle soak:** No cues for **15+ minutes**; session should remain up (no repeated disconnects / reboots).

### Gatekeeper / unsigned builds

If `identity` is null in `package.json` **mac** section, builds may be **ad-hoc** signed. On prod Macs, operators may need **right-click → Open** the first time, or IT approval per your standard process.

---

## Quick pass / fail checklist

- [ ] Cable pinout verified against **DSAN + WaveShare** docs (not guessed).
- [ ] WaveShare serial **8N1** baud matches PerfectCue output (adjust if raw data is garbage).
- [ ] TCP **Client** → controller IP + **same** port as PerfectCue row.
- [ ] App row set to **WaveShare**; settings saved.
- [ ] Keep-alive / no-data restart aligned with long idle (section 3).
- [ ] **`yarn build:mac`** artifact tested on target CPU arch; PerfectCue logs show **`(waveshare)`** on connect.
