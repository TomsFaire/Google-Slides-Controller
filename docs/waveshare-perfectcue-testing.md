# WaveShare + PerfectCue — testing guide

This guide covers bench validation of the Google Slides Controller PerfectCue path with a **WaveShare** serial–Ethernet adapter. It assumes IP/network parameters are already correct on the converter.

**Related:** [`waveshare-perfectcue-setup.md`](waveshare-perfectcue-setup.md) (network and field checklist).

---

## Safety and pinouts (read first)

1. **PerfectCue RJ45 ports are model-specific.** Some DSAN ports carry **inter-unit cue data and/or power** (see DSAN PerfectCue / cue-light documentation). **Do not** assume “standard Ethernet” T568A/B pinout or plug an unknown homemade cable into audio gear or a switch.
2. **Confirm the exact jack you use.** Your integration uses the **RS485 / network extender path** as discussed with DSAN—verify in the **manual for your PerfectCue model** which RJ45 pins are **RS485 A/B** (and whether **GND** or **shield** is required). If documentation is unclear, contact DSAN support with your model number before applying DC or long cables.
3. **WaveShare connectors.** On **RS232/485/422 TO POE ETH (B)** the **DB9 is RS-232**—do **not** use it for PerfectCue RS485. Use the **RS422/RS485** screw terminals (see [WaveShare wiki](https://www.waveshare.com/wiki/RS232/485/422_TO_POE_ETH_(B))). For **RS485**, official wiring is **positive → TA**, **negative → TB**; **RB** / extra **TA** are for **RS422**—leave unused for a simple RS485 link. Block order on the panel may read **RB, TA, TA, TB, PE**, then **GND**, **VCC**—the **last two** are **DC power in** only—see **Power** below.

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

### How many wires? (read this first)

At the WaveShare **RS422/485** row you land **exactly three conductors** for this job:

1. **One wire** = DSAN **Data+** (RJ45 pins **3** and **6** strapped together **at the RJ45 plug** so they share **one** copper leaving toward WaveShare).
2. **One wire** = DSAN **Data−** (RJ45 pins **7** and **8** strapped together the same way).
3. **One wire** = DSAN **ground** (RJ45 pins **4** and **5** strapped together).

Pins **1** and **2** do **not** connect to anything at the WaveShare end.

**Important:** **Data+** goes to **exactly one** screw on the block. **Data−** goes to **exactly one other** screw. You do **not** fan out one signal onto multiple terminals.

### Official RS485 pairing ([RS232/485/422 TO POE ETH (B)](https://www.waveshare.com/wiki/RS232/485/422_TO_POE_ETH_(B)))

WaveShare’s wiki (*Hardware Connection*) states for **RS485**:

> … connect the **positive pole to the TA**, the **negative pole to the TB** …

So for **this product family**, **two-wire RS485** uses **only TA and TB** for the differential pair:

| DSAN side (after strapping in the RJ45) | WaveShare terminal |
|----------------------------------------|--------------------|
| **Data+** (RJ45 **3** + **6** → one wire) | **TA** |
| **Data−** (RJ45 **7** + **8** → one wire) | **TB** |
| **GND** (RJ45 **4** + **5** → one wire) | **PE** (protective earth / shield—follow manual if ref differs) |

**RB** and the **extra TA** on the silkscreen are used for **RS422** (and related modes)—for simple **half-duplex RS485** to PerfectCue they stay **empty** unless your manual says otherwise.

If bytes look wrong after baud is correct, **swap TA and TB once** (exchange Data+ and Data−).

Other WaveShare SKUs may label differently—always cross-check the wiki page for **your** model string.

### End-to-end checklist

| Step | What to do |
|------|------------|
| DSAN plug | Strap **3+6** → **one** conductor (**Data+**). Strap **7+8** → **one** (**Data−**). Strap **4+5** → **one** (**GND**). Leave **1–2** unconnected at WaveShare. |
| WaveShare ([wiki](https://www.waveshare.com/wiki/RS232/485/422_TO_POE_ETH_(B))) | **Data+** → **TA**. **Data−** → **TB**. **GND** → **PE** (unless manual differs). |
| Wrong data? | Swap **TA** ↔ **TB** (swap the two differential wires only). |

### Diagram (RS232/485/422 TO POE ETH (B) — RS485)

```
RJ45 (strapped)                          WaveShare RS485

  pin 3 ─┬── ONE wire "Data+" ──────────► TA
  pin 6 ─┘

  pin 7 ─┬── ONE wire "Data−" ──────────► TB
  pin 8 ─┘

  pin 4 ─┬── ONE wire "GND" ─────────────► PE
  pin 5 ─┘

  pin 1,2   (not connected at WaveShare)

  RB, extra TA    (leave empty for RS485-only link)
```

### Troubleshooting: PerfectCue shuts off when the cable is plugged in

That behavior usually means the DSAN side sees a **short on the 12 V path** (RJ45 **pins 1–2**) or **12 V shorted to ground**—often **before** any RS485 signaling matters.

1. **Verify RJ45 pin numbers with a multimeter**, not only wire colors. **T568A vs T568B** swaps which **color** lands on pin 3 vs pin 1. **Pins 1–2** are often **solid orange / white-orange** (T568B)—those pairs must go **nowhere** on the WaveShare pigtail. If orange conductors accidentally touch **TA**, **TB**, **PE**, or each other at the crimp, you can short **12 V** and trip protection.

2. **Inspect the RJ45 crimp**: stray copper strands bridging adjacent pins are a common cause.

3. **First retest with only two wires** (no PE): DSAN **3+6 → TA**, **7+8 → TB** only; leave **4–5** and **PE** disconnected temporarily. If it **still** dies, the fault is almost certainly **wrong pins on the plug** or **12 V involved**—not the PE link alone.

4. **Terminal positions:** Count screws from the same reference as the silkscreen (**RB**, **TA**, **TA**, **TB**, **PE**). **Data+** must land on **a** terminal marked **TA** (use the **first TA** next to **RB** if there are two); **Data−** on **TB**; **PE** last in that group. If **Data+** was wired to the wrong screw (e.g. **RB**), behavior depends on the board—fix per wiki (**TA** / **TB**).

5. **Do not power-loop** through this cable: this harness must **not** carry DSAN **12 V** to WaveShare **VCC/GND**.

---

## 1. Building the cable (RJ45 on PerfectCue → WaveShare RS485 block)

### What you need

- Cable suitable for **RS485 + ground** (and optional shield to **PE** if your manual recommends it).
- **Do not** connect RJ45 **pins 1–2** to WaveShare unless DSAN and WaveShare documentation explicitly define a safe shared-power arrangement (default: **leave open**).
- **WaveShare side:** terminate only on the **RS422/RS485** terminals (**RB / TA / TA / TB / PE**), not the RS-232 DB9; **GND/VCC** at the right end of the block are **device power**, not RS485 data.
- **RS232/485/422 TO POE ETH (B):** **Data+ → TA**, **Data− → TB**, **GND → PE** ([wiki](https://www.waveshare.com/wiki/RS232/485/422_TO_POE_ETH_(B))).

### Wiring steps (procedure)

1. Follow **[Signal-only cable pairing chart](#signal-only-cable-pairing-chart-dsan-rj45--waveshare-block)** above for pin-to-terminal mapping.
2. Connect **Data+** and **Data−** to the WaveShare terminals your manual assigns to **RS485 A/B** (half-duplex). Connect **Ground** (RJ45 **4–5**) per the chart (**PE** / shield rules).
3. **Do not** connect RJ45 **12 V** (pins 1–2) to **VCC/GND** unless a qualified doc says so (default: **floating / not wired** at WaveShare).
4. If raw data looks wrong but polarity is unsure, **swap** the two differential wires once (RS485 A/B swap).
5. **First power-on:** Power WaveShare normally; power PerfectCue normally; then check the app debug console for `[PerfectCue] raw:` when pressing forward/back.

If your unit is **not** RS232/485/422 TO POE ETH (B), confirm **TA/TB** roles on the wiki page for **your** model.

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

Goal: the converter must stay connected through idle periods **without** requiring you to poll manually. App presets live in `src/perfectcue-adapter-presets.js` (**WaveShare:** `0xFF` ping about every **45 s**, Node read-idle **120 s**).

### Which setting is “TCP keep-alive”? (you may not see it in the browser)

On many WaveShare units the **simple web page** (`http://<ip>/…`) shows **Reconnect-time** and **No-data restart**, but **does not show** a separate **TCP keep-alive interval** field. That parameter is often available only in **Vircom**:

1. Install/open **Vircom** (WaveShare’s Windows config tool—see [RS232/485/422 TO POE ETH (B) wiki](https://www.waveshare.com/wiki/RS232/485/422_TO_POE_ETH_(B)) → Software).
2. **Device Management** → select your device → open **Device Settings** (double-click).
3. Go to **Advanced Settings** (or **More Advanced Settings…** if present).
4. Look for **Keep Alive Time** (seconds)—this is the firmware field that controls how often the module treats the TCP link as “still alive” from its side. Set it to **≥ 60 s** (e.g. **60**) so it does **not** tear down the session *before* the app’s next **`0xFF`** ping at **45 s**. If it were **30 s** while the app pings every **45 s**, you could see idle drops.

If you **only** use the browser and **never** see “Keep Alive Time,” that’s normal for the stripped-down UI—use Vircom on a PC on the same LAN, or skip tuning this if your link **already stays up** (the app’s **`0xFF`** traffic often keeps the session alive anyway).

### Settings you *do* usually see (web or Vircom)

- **Reconnect time** (~**12 s**): How soon the **TCP client** retries **after** a disconnect—not the same as keep-alive. Helps after the controller app restarts.
- **No-data restart** / **Restart for no data**: Prefer **Disabled** so the bridge does not reset during long idle periods with no cues.
- **Keep Alive Time** (often **Vircom → Advanced Settings** only): Target **≥ 60 s** so the firmware does not drop TCP *before* the app’s **45 s** `0xFF` ping; avoid **30 s** if the app pings every **45 s**.

After changes, **Modify Setting** / **Submit** and reboot the device if prompted.

### Vircom shortcut

**Vircom:** Device → **Advanced Settings** → **Keep Alive Time**, **Reconnect Time**, turn off **Restart for no data** when possible → **Modify Setting** → restart if asked.

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

- [ ] Cable: **three** wires to RS485 block (**Data+**, **Data−**, **GND→PE**) to **exactly two** RS485 screws + **PE** per manual—not split across every TA/TB/RB; DB9 unused; RJ45 **12 V** not tied to **VCC**.
- [ ] WaveShare serial **8N1** baud matches PerfectCue output (adjust if raw data is garbage).
- [ ] TCP **Client** → controller IP + **same** port as PerfectCue row.
- [ ] App row set to **WaveShare**; settings saved.
- [ ] Keep-alive / no-data restart aligned with long idle (section 3).
- [ ] **`yarn build:mac`** artifact tested on target CPU arch; PerfectCue logs show **`(waveshare)`** on connect.
