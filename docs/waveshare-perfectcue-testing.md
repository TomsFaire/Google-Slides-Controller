# WaveShare + PerfectCue — testing guide

This guide covers bench validation of the Google Slides Controller PerfectCue path with a **WaveShare** serial–Ethernet adapter. It assumes IP/network parameters are already correct on the converter.

**Related:** [`waveshare-perfectcue-setup.md`](waveshare-perfectcue-setup.md) (network and field checklist).

---

## Safety and pinouts (read first)

1. **PerfectCue RJ45 ports are model-specific.** Some DSAN ports carry **inter-unit cue data and/or power** (see DSAN PerfectCue / cue-light documentation). **Do not** assume “standard Ethernet” T568A/B pinout or plug an unknown homemade cable into audio gear or a switch.
2. **Confirm the exact jack you use.** Your integration uses the **RS485 / network extender path** as discussed with DSAN—verify in the **manual for your PerfectCue model** which RJ45 pins are **RS485 A/B** (and whether **GND** or **shield** is required). If documentation is unclear, contact DSAN support with your model number before applying DC or long cables.
3. **WaveShare connectors.** On this hardware the **DE-9 (DB9) is labeled RS-232**—do **not** use it for PerfectCue RS485. Use the **left-hand screw terminals** labeled **RS422/RS485** (silkscreen order from the DIN rail side: **RB, TA, TA, TB, PE, GND**, then **VCC**). Land RS485 + reference/shield per the WaveShare manual’s meaning of **TA/TB/RB** for **half-duplex RS485** (ferrules recommended). The **last two** terminals (**GND**, **VCC**) are the **DC power input** for the module—see **Power (VCC/GND) vs DSAN 12 V** below.

---

## Power (VCC/GND) vs DSAN 12 V on RJ45

**What WaveShare `VCC` / `GND` are:** Those two terminals are **DC power supplied into the WaveShare** (**6–36 V** per silkscreen), same role as the barrel jack / PoE options—**inbound** power for **this** converter. They are **not** a 12 V **output** designed to power a PerfectCue.

**Can WaveShare power the PerfectCue?** **No**, not through these terminals in normal use. Do **not** wire DSAN RJ45 **pins 1–2 (12 V)** into WaveShare **VCC/GND** hoping to “feed” one box from the other without a written compatibility check—polarity, current, and regulation differ from consumer interconnects, and **mis-wiring can damage hardware**.

**Should DSAN 12 V be connected for this link?** For a **signal-only** bridge to the WaveShare, **leave RJ45 pins 1–2 unconnected** at the WaveShare end. Power **PerfectCue** the usual way (its own adapter or DSAN **RJ45 power between DSAN devices**). Power **WaveShare** with **PoE**, **barrel**, or **VCC/GND** as its manual specifies.

**RS232 vs RS485 voltage:** DSAN’s note about “voltage differences” vs RS232 is why we use the **RS485 block**, not the RS-232 DB9.

---

## DSAN RJ45 pinout (support email — cue-light RS485)

RJ45 numbering is **T568-A/B standard pin numbers** (tab down, contacts numbered 1–8 left to right). Confirm with DSAN for your **exact** model before soldering.

- **Pins 1, 2:** **12 V** (inter-unit power — usually **not** wired to WaveShare; see **Power** section).
- **Pins 3 + 6:** **Data+** (differential; often tied together at DSAN).
- **Pins 7, 8:** **Data−**.
- **Pins 4, 5:** **Ground**.

**Serial bytes** on that link (match Google Slides Controller parser): **0x0F** forward, **0x1F** back; DSAN also lists **0x2F** blank on, **0x3F** blank off (not used by the app unless extended).

---

## 1. Building the cable (RJ45 on PerfectCue → WaveShare RS485 block)

### What you need

- Cable suitable for **RS485 + ground** (and optional shield to **PE** if your manual recommends it).
- **Do not** connect RJ45 **pins 1–2** to WaveShare unless DSAN and WaveShare documentation explicitly define a safe shared-power arrangement (default: **leave open**).
- **WaveShare side:** terminate only on the **RS422/RS485** terminals (**RB / TA / TA / TB / PE**), not the RS-232 DB9; **GND/VCC** at the right end of the block are **device power**, not RS485 data.
- Map **Data+ / Data− / Ground** from the RJ45 to the correct **RS485 differential pair** on **TA/TB/RB** per **WaveShare’s manual** for **RS485 two-wire** mode (labels vary by firmware revision).

### Wiring steps (procedure)

1. From the table above, identify **Data+**, **Data−**, and **Ground** on the RJ45 pigtail.
2. Connect **Data+** and **Data−** to the WaveShare terminals your manual assigns to **RS485 A/B** (half-duplex). Connect **Ground** (RJ45 4–5) to **GND** on the signal block if the manual ties RS485 reference there—use **PE** only per shield/earth guidance in the manual.
3. **Do not** connect RJ45 **12 V** (pins 1–2) to **VCC/GND** unless a qualified doc says so (default: **floating / not wired** at WaveShare).
4. If raw data looks wrong but polarity is unsure, **swap** the two differential wires once (RS485 A/B swap).
5. **First power-on:** Power WaveShare normally; power PerfectCue normally; then check the app debug console for `[PerfectCue] raw:` when pressing forward/back.

If anything is ambiguous, stop and confirm **TA/TB/RB** roles for **RS485** in the WaveShare wiki/manual for your exact SKU—RS422 labeling can differ from two-wire RS485.

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

- [ ] Cable pinout verified against **DSAN + WaveShare** docs (not guessed); RS485 on **RB/TA/TB/PE** block only (DB9 is RS-232); RJ45 **12 V** pins **not** tied to WaveShare **VCC** unless documented.
- [ ] WaveShare serial **8N1** baud matches PerfectCue output (adjust if raw data is garbage).
- [ ] TCP **Client** → controller IP + **same** port as PerfectCue row.
- [ ] App row set to **WaveShare**; settings saved.
- [ ] Keep-alive / no-data restart aligned with long idle (section 3).
- [ ] **`yarn build:mac`** artifact tested on target CPU arch; PerfectCue logs show **`(waveshare)`** on connect.
