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

RJ45 numbering uses **standard Ethernet jack numbering**: hold the plug **latch tab facing down / away from you**, pins **1–8** left to right (metal contacts visible toward you). Confirm with DSAN for your **exact** model before soldering.

- **Pins 1, 2:** **12 V** (inter-unit power — usually **not** wired to WaveShare; see **Power** section).
- **Pins 3 + 6:** **Data+** (same polarity; often tied together at DSAN—you may wire **one** conductor or both).
- **Pins 7, 8:** **Data−** (same polarity; tie together at your pigtail if you use both).
- **Pins 4, 5:** **Ground**.

**Serial bytes** on that link (match Google Slides Controller parser): **0x0F** forward, **0x1F** back; DSAN also lists **0x2F** blank on, **0x3F** blank off (not used by the app unless extended).

---

## Signal-only cable pairing chart (DSAN RJ45 ↔ WaveShare block)

This is the **signal-only** link: **no** RJ45 **12 V** (pins **1–2**) at the WaveShare end. Power each box separately (PerfectCue via DSAN/adapters; WaveShare via PoE / barrel / **VCC+GND**).

WaveShare RS422/485 terminals (left → right from DIN side): **RB, TA, TA, TB, PE**, then **DC: GND, VCC**. Only **RB … PE** carry RS485/shield; **GND/VCC** are **DC power in** for the module—not RS485 data.

**PE vs the DC `GND` next to `VCC` (not the same job):**

- **`GND` beside `VCC`:** This is the **DC power return** (negative) for the **6–36 V** input that **powers the WaveShare** (with `VCC` = positive). Use it only for your **barrel / screw-terminal power supply** leads—not as a stand-in for RS485 “signal ground” unless the **manual or schematic** says the interface reference is tied there.
- **`PE`:** **Protective earth**—safety / chassis / **cable shield** drain. Field RS485 wiring often brings the cable shield (and sometimes the partner’s signal reference) to **PE** per manufacturer guidance.

So: **`GND` = power supply ground reference**; **`PE` = protective earth / shield path** (and often where DSAN **RJ45 4–5** land **if** the manual agrees). On some boards **PE** and internal signal reference are related; on others isolation separates them—**follow WaveShare’s doc** and don’t assume **`GND` (power)** and **`PE`** are interchangeable without checking.

### End-to-end connections

| DSAN RJ45 pin(s) | Name (DSAN) | Connect to WaveShare terminal | Notes |
|-------------------|-------------|-------------------------------|--------|
| **1** | 12 V | **(none)** | Leave **open** at WaveShare—do not tie to **VCC/GND**. |
| **2** | 12 V | **(none)** | Same as pin 1. |
| **3** and **6** | **Data+** | **RS485 “A” / non-inverting / D+** | On many WaveShare boards this is one of **TA** or **TB**—read the manual for **RS485 two-wire** which screw is **A** (+). Tie RJ45 **3** and **6** together at your cable if you run one wire. |
| **7** and **8** | **Data−** | **RS485 “B” / inverting / D−** | Usually the other of **TA/TB** or **RB**—manual will label **B** (−). Tie RJ45 **7** and **8** together at your cable if you run one wire. |
| **4** and **5** | **Ground** | **PE** (and/or signal reference per manual) | Tie RJ45 **4** and **5** together at your cable. **PE** is for protective earth / shield; follow WaveShare doc for whether signal ref ties here vs another terminal. |

If you have **no bytes** or garbage after baud checks, **swap** only the two differential wires (swap which WaveShare terminals get **Data+** vs **Data−**).

### Which TA vs TB vs RB?

Silkscreen varies by SKU and **RS422 vs RS485** mode. Your block shows **RB**, **TA**, **TA**, **TB**, **PE**:

- Use the WaveShare wiki/manual for **RS232 RS485 TO POE ETH (B)** (or your exact model) to see which screws are **RS485 A** and **RS485 B** in **half-duplex RS485**.
- The **duplicate “TA”** may be two adjacent pads both marked TA, or TX/RX legs—**confirm** whether one is **no-connect** in RS485 mode.

Until the manual is checked, treat the table above as **nets**: **Data+**, **Data−**, **GND/shield**—assign them to the physical screws your documentation calls **A**, **B**, and earth.

### One-line summary

```
RJ45 3+6 (Data+)  ──►  WaveShare RS485 A / (+) terminal (TA or TB per manual)
RJ45 7+8 (Data−)  ──►  WaveShare RS485 B / (−) terminal (RB or TB per manual)
RJ45 4+5 (GND)    ──►  PE (and reference rules per manual)

RJ45 1+2 (12 V)   ──►  not connected at WaveShare
```

---

## 1. Building the cable (RJ45 on PerfectCue → WaveShare RS485 block)

### What you need

- Cable suitable for **RS485 + ground** (and optional shield to **PE** if your manual recommends it).
- **Do not** connect RJ45 **pins 1–2** to WaveShare unless DSAN and WaveShare documentation explicitly define a safe shared-power arrangement (default: **leave open**).
- **WaveShare side:** terminate only on the **RS422/RS485** terminals (**RB / TA / TA / TB / PE**), not the RS-232 DB9; **GND/VCC** at the right end of the block are **device power**, not RS485 data.
- Map **Data+ / Data− / Ground** from the RJ45 to the correct **RS485 differential pair** on **TA/TB/RB** per **WaveShare’s manual** for **RS485 two-wire** mode (labels vary by firmware revision).

### Wiring steps (procedure)

1. Follow **[Signal-only cable pairing chart](#signal-only-cable-pairing-chart-dsan-rj45--waveshare-block)** above for pin-to-terminal mapping.
2. Connect **Data+** and **Data−** to the WaveShare terminals your manual assigns to **RS485 A/B** (half-duplex). Connect **Ground** (RJ45 **4–5**) per the chart (**PE** / shield rules).
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
