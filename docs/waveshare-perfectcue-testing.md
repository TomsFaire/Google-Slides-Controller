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
- **Pins 3 + 6:** **Data+** (DSAN may run both to the same net **inside** their gear—details below on **not** jumpering them at **your** RJ45).
- **Pins 7, 8:** **Data−**.
- **Pins 4, 5:** **Ground**.

**Serial bytes** observed on this link at 9600 baud with WaveShare RS232/485/422 TO POE ETH (B):

| Button | Clean byte | RS485 noise variant | App action |
|--------|-----------|---------------------|------------|
| Forward / Next | `0x0c` | `0x8c` | next-slide |
| Back / Previous | `0x08` | `0x88` | previous-slide |
| Blackout | `0x04` | `0x84` | *(stub — not yet dispatched)* |

The app parser masks the high bit (`byte & 0x7f`) before lookup, so both the clean byte and its noisy variant map to the same command. The `0x80` high bit is set intermittently by framing noise on unterminated RS485; adding a **120 Ω termination resistor** across TA/TB should suppress it.

---

## Signal-only cable pairing chart (DSAN RJ45 ↔ WaveShare block)

This is the **signal-only** link: **no** RJ45 **12 V** (pins **1–2**) at the WaveShare end. Power each box separately (PerfectCue via DSAN/adapters; WaveShare via PoE / barrel / **VCC+GND**).

WaveShare RS422/485 terminals (left → right from DIN side): **RB, TA, TA, TB, PE**, then **DC: GND, VCC**. Only **RB … PE** carry RS485/shield; **GND/VCC** are **DC power in** for the module—not RS485 data.

**PE vs the DC `GND` next to `VCC` (not the same job):**

- **`GND` beside `VCC`:** This is the **DC power return** (negative) for the **6–36 V** input that **powers the WaveShare** (with `VCC` = positive). Use it only for your **barrel / screw-terminal power supply** leads—not as a stand-in for RS485 “signal ground” unless the **manual or schematic** says the interface reference is tied there.
- **`PE`:** **Protective earth**—safety / chassis / **cable shield** drain. Field RS485 wiring often brings the cable shield (and sometimes the partner’s signal reference) to **PE** per manufacturer guidance.

So: **`GND` = power supply ground reference**; **`PE` = protective earth / shield path** (and often where DSAN **RJ45 4–5** land **if** the manual agrees). On some boards **PE** and internal signal reference are related; on others isolation separates them—**follow WaveShare’s doc** and don’t assume **`GND` (power)** and **`PE`** are interchangeable without checking.

### Do **not** strap 3+6, 7+8, or 4+5 together at the RJ45 (unless DSAN says so)

DSAN’s email describes **parallel** roles (**3** and **6** both “Data+”, etc.). That does **not** mean **your** pigtail must **electrically short** those pins inside the RJ45 plug.

On some PerfectCue / network-extender jacks, **externally jumpering** pin **3** to **6**, **7** to **8**, or **4** to **5** can:

- Pull together PCB nets that are **not** meant to be tied **at the connector**, or  
- Present an abnormal load that trips **overcurrent / foldback** (unit **powers off** when plugged in).

Field fix that often works: use **one pin per function**—**no straps** between pins at the plug:

| Function | Prefer **one** RJ45 pin first | WaveShare |
|----------|-------------------------------|-----------|
| Data− | **Pin 7 only** → TA (leave **pin 8** open on pigtail) | **TA** |
| Data+ | **Pin 3 only** → TB (leave **pin 6** **not** jumpered to pin 3) | **TB** |
| Ground | **Pin 4 only** → PE (leave **pin 5** open) — add pin 5 later only if needed | **PE** |

> **Field-confirmed swap:** Despite the WaveShare wiki labelling TA as positive and TB as negative, field testing with DSAN PerfectCue RS485 requires **pin 7 (Data−) → TA** and **pin 3 (Data+) → TB**. If you see no RS485 activity, swap TA ↔ TB.

If RS485 is unreliable with single-pin taps, ask DSAN whether **your** jack allows paralleling **3+6** / **7+8** / **4+5** **externally**—then try straps **only** after written OK.

**Important:** **Data+** / **Data−** each go to **exactly one** WaveShare screw (**TA** / **TB**). Pins **1–2** stay **unconnected**.

### Official RS485 pairing ([RS232/485/422 TO POE ETH (B)](https://www.waveshare.com/wiki/RS232/485/422_TO_POE_ETH_(B)))

WaveShare’s wiki states for **RS485**: **positive → TA**, **negative → TB**. However, field testing with this DSAN PerfectCue unit requires the opposite — **pin 7 → TA**, **pin 3 → TB**. The parser and byte values are unaffected; only the physical screw assignment is swapped.

**RB** and the **extra TA** stay empty for simple half-duplex RS485.

### End-to-end checklist

| Step | What to do |
|------|------------|
| DSAN plug | **Pin 7** (Brown/White) → **TA**, **pin 3** (Green/White) → **TB**, **pin 4** (Blue) → **PE**. Pins **1–2** unused. No 3+6 / 7+8 / 4+5 straps. |
| WaveShare | **TA** / **TB** / **PE** per above (note swap vs wiki). |
| Wrong data? | Swap **TA** ↔ **TB**. |

### Diagram (single-pin taps — recommended if straps trip protection)

```
RJ45                                    WaveShare RS485

  pin 7 only (Brown/White) ──────────► TA     (Data−; pin 8 not jumpered to 7)
  pin 3 only (Green/White) ──────────► TB     (Data+; pin 6 not jumpered to 3)
  pin 4 only (Blue)        ──────────► PE     (pin 5 not jumpered to 4)

  pins 1,2   not connected            100Ω across TA↔TB (termination)

  pin 6,8,5  left unterminated on pigtail unless DSAN approves straps
```

> **Note:** This is the opposite of the WaveShare wiki convention (which shows positive→TA). Swap confirmed by field testing — if no RS485 activity, try swapping TA and TB.

### Troubleshooting: PerfectCue shuts off when the cable is plugged in

**If shutdown happens even with WaveShare disconnected** and only your RJ45 pigtail inserted, the fault is **entirely** at the **RJ45**—almost always **straps or wrong pins**.

1. **Identify which strap causes it** (you reported one of **3+6**, **7+8**, **4+5**): **remove that strap first**. Preferred fix: **stop strapping**—use **single-pin** taps (pin **3**, **7**, **4** only) as in the diagram above.

2. **Verify pin numbers** with a meter from **each gold contact** to its wire—**tab-down** pin **1…8**—don’t rely on **ethernet pair colors** until verified (**T568A vs B** swaps colors vs pin numbers).

3. **Orange pair** (often pins **1–2**, **12 V** on DSAN cue wiring) must have **no** copper reaching TA/TB/PE and **no** stray strands bridging into pins **3–6**.

4. **Bad crimp**: strands shorting **adjacent** pins (e.g. **2** touching **3**) put **12 V** into RS485—same shutdown symptom.

5. **Ask DSAN** whether **your** RJ45 port matches the cue-light pinout email **exactly** for **network extender** vs **cue interconnect**—if paralleling **3+6** is invalid on **your** PCB, single-pin pickup is correct.

---

## 1. Building the cable (RJ45 on PerfectCue → WaveShare RS485 block)

### What you need

- Cable suitable for **RS485 + ground** (and optional shield to **PE** if your manual recommends it).
- **Do not** connect RJ45 **pins 1–2** to WaveShare unless DSAN and WaveShare documentation explicitly define a safe shared-power arrangement (default: **leave open**).
- **WaveShare side:** terminate only on the **RS422/RS485** terminals (**RB / TA / TA / TB / PE**), not the RS-232 DB9; **GND/VCC** at the right end of the block are **device power**, not RS485 data.
- **RS232/485/422 TO POE ETH (B):** **Data+ → TA**, **Data− → TB**, **GND → PE** ([wiki](https://www.waveshare.com/wiki/RS232/485/422_TO_POE_ETH_(B))).

### Wiring steps (procedure)

1. Follow **[Signal-only cable pairing chart](#signal-only-cable-pairing-chart-dsan-rj45--waveshare-block)** — prefer **single-pin** DSAN taps (**3**, **7**, **4**) without strapping pairs at the RJ45 unless DSAN confirms.
2. Connect **TA** / **TB** / **PE** per wiki; do **not** fan multiple RJ45 pins onto one screw without DSAN approval.
3. **Do not** connect RJ45 **12 V** (pins 1–2) to **VCC/GND** unless a qualified doc says so (default: **floating / not wired** at WaveShare).
4. If raw data looks wrong but polarity is unsure, **swap** the two differential wires once (RS485 A/B swap).
5. **First power-on:** Power WaveShare normally; power PerfectCue normally; then check the app debug console for `[PerfectCue] raw:` when pressing forward/back.

If your unit is **not** RS232/485/422 TO POE ETH (B), confirm **TA/TB** roles on the wiki page for **your** model.

---

## 2. WaveShare serial settings (after IP is set)

Configure via **web UI** (`http://<device-ip>/`) and/or **Vircom** so they match **what PerfectCue outputs on that RS485 link** and match the **Google Slides Controller** listener.

### Typical values (confirm against your hardware)

- **Baud rate:** **9600** — confirmed in field testing with WaveShare RS232/485/422 TO POE ETH (B) and DSAN PerfectCue RS485 output. At 115200 the WaveShare samples each PerfectCue bit ~12× too fast, producing all-zero frames. If your PerfectCue model differs, try 115200 only if 9600 produces garbage.
- **Data bits:** **8**
- **Parity:** **None**
- **Stop bits:** **1**
- **Flow control:** **None**

Apply / submit settings and **restart the device** if the UI requires it.

### Work mode (TCP Client vs TCP Server / UDP)

Google Slides Controller opens a **TCP server** on each PerfectCue listener port (`createServer` in [`src/perfectcue-server.js`](../src/perfectcue-server.js)). The WaveShare must **connect out** to that machine—use **TCP Client** with **Destination IP** = controller host and **Destination Port** = the same listener port (e.g. **8899**).

- **TCP Client:** **Yes** — WaveShare connects **out** to the PC’s PerfectCue listener port.
- **TCP Server:** **No** for the stock setup — that would make WaveShare the listener and require the PC to dial **in** to it (opposite of how [`perfectcue-server.js`](../src/perfectcue-server.js) works).
- **UDP** / **UDP Group:** **No** — the app listens on **TCP** only.

### Transparent serial (required for PerfectCue)

PerfectCue sends **arbitrary binary bytes** over RS485 (not Modbus). For raw bytes to appear on TCP unchanged:

- In VirCom **Device Settings** → **Function of the device**, **uncheck** **Modbus TCP to RTU**. With Modbus enabled, the gateway expects **Modbus TCP** on Ethernet and **Modbus RTU** framing on serial—PerfectCue traffic will **not** pass through as transparent payloads (you may see “TCP connected” in VirCom but **no useful bytes** with **`nc`** or the Slides app).
- Set **Transfer Protocol** / conversion mode to **None** (transparent) wherever it appears.
- The web UI note (*Multi-host is always enabled when Protocol is Modbus TCP to RTU*) is another hint to **leave Modbus off** for this use case.

After disabling Modbus, **Modify Setting** / reboot the WaveShare and retest **`nc`** / PerfectCue buttons.

### If **Modbus TCP to RTU** is checked and **grayed out**

VirCom locks that checkbox when **another setting forces gateway / Modbus behavior**. Clear the dependency first—then Modbus can be turned off.

1. **Advanced Settings / protocol fields** — Look for **Conversion protocol**, **Transfer Protocol**, **Gateway type**, **Protocol**, or **Serial port mode** set to **Modbus** (any variant). Set to **None**, **Transparent**, **Direct**, or **TCP/IP transparent transmission** (exact wording varies). Apply / reboot, then reopen **Function of the device** and try Modbus again.

2. **RS485 Multi-Host / Modbus Gateway** dialog — If **Modbus Gateway Type** is set to **Simple Modbus TCP to RTU**, **Pre-configurable Modbus GW**, etc., change it to an option that **disables** gateway behavior, or close the dialog after clearing **Enable** options—depending on firmware. Some builds only unlock Modbus after gateway type is cleared on the **web UI**.

3. **Web browser config** (`http://<device-ip>/`) — Often the **same** Modbus flags appear there; sometimes you can set **transparent** mode in the browser when VirCom grayed it out. Submit and **restart device**.

4. **Other Function checkboxes** — Rarely **REAL_COM**, **cloud**, or **MQTT/JSON** modes interact with protocol selection; try disabling nonessential features one at a time if docs suggest a conflict.

5. **Last resort** — **Factory reset** (long **RESET** per [wiki](https://www.waveshare.com/wiki/RS232/485/422_TO_POE_ETH_(B))), then configure **only**: network, **TCP Client**, destination IP/port, serial **8N1** baud, **transparent** protocol—**do not** enable Modbus or Modbus gateway during initial setup.

### Blackout button — current status

The blackout button sends byte `0x04` (or `0x84` with RS485 high-bit noise). The app **recognises** it — the parser returns `'blackout'` — but does **not yet dispatch** any action. The debug console will show:

```
[PerfectCue] blackout cue received (not yet implemented)
```

**Future implementation:** pressing blackout should send a `B` keypress to the Google Slides presentation window, which toggles the screen black (same as pressing **B** in Chrome during a slideshow). This will be wired up in `main.js` once the signal is clean (termination resistor fitted and `0x04` confirmed reliable).

**Known issue (unterminated RS485):** Without the termination resistor, `0x04` sometimes arrives as `0x08` (one bit flipped by a reflection), which the app treats as a previous-slide command. Fitting the 120 Ω resistor across TA/TB is the fix — do not attempt to distinguish blackout from back in software until the hardware signal is stable.

### App-side listener

- In **Settings → PerfectCue**, set the row’s **Converter** dropdown to **WaveShare** for longer keep-alive presets (vs **DSAN**).
- **Save** PerfectCue settings.
- Ensure the WaveShare is still in **TCP Client** mode to the controller’s IP and the **same TCP port** as that row (default **8899**).

### Validating serial data reaches WaveShare and the PC

You care about **three links**: RS485 **PerfectCue → WaveShare**, **WaveShare serial → TCP**, **TCP → controller app**.

1. **TCP session up (WaveShare ↔ PC)**  
   In **Vircom → Device Management**, confirm the device shows an **established TCP client** connection to your controller **IP:PerfectCue port**. In the **web UI**, look for connection / status fields if present. No TCP link → fix IP, port, firewall, or **TCP Client** mode before debugging RS485.

2. **Bytes moving at the serial port (WaveShare)**  
   Firmware varies: some builds show **RX/TX counters**, a **serial traffic / debug** view, or **LED activity** on RS485—check the [wiki](https://www.waveshare.com/wiki/RS232/485/422_TO_POE_ETH_(B)) and Vircom help for **your** revision. If nothing is exposed, rely on steps 3–4.

3. **End-to-end with Google Slides Controller (strongest check)**  
   Enable **PerfectCue**, open the **debug console**, press **forward/back** on the remote. You want **`client connected (waveshare)`** and lines like **`[PerfectCue] raw:`** with hex (e.g. **`0c`** / **`08`** for next/previous, or high-bit variants **`8c`** / **`88`** if the RS485 line is unterminated). That proves bytes left PerfectCue, entered WaveShare on RS485, were forwarded over TCP, and reached the app. **All-zero frames** mean baud mismatch (check 9600); **no bytes at all** means Modbus is still on or transparent mode isn't set.

4. **Optional: raw TCP without the full app**  
   On the PC/Mac running **`nc`**, listen on the **same port** WaveShare uses as **Destination Port** (e.g. **`nc -lk 4196`** or **`nc -l 0.0.0.0 4196`** so you bind IPv4). Press forward/back on the remote—you should see **binary bytes**. If the TCP session shows up but **no bytes**, fix **transparent mode** (above) and **serial baud** (try **115200** if **9600** shows nothing). Stop **`nc`** before the real app listens on that port.

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

- [ ] Cable: **pin 7** (Brown/White) → **TA**, **pin 3** (Green/White) → **TB**, **pin 4** (Blue) → **PE**; no RJ45 straps; DB9 unused; pins 1–2 unused. If no RS485 activity, swap TA↔TB.
- [ ] WaveShare serial **9600 8N1** (confirmed for DSAN PerfectCue RS485; all-zero frames = wrong baud).
- [ ] **120 Ω termination resistor** across TA/TB recommended to suppress RS485 high-bit framing noise (symptoms: `0x8c` instead of `0x0c`, or blackout/back indistinguishable).
- [ ] TCP **Client** → controller IP + **same** port as PerfectCue row.
- [ ] App row set to **WaveShare**; settings saved.
- [ ] Keep-alive / no-data restart aligned with long idle (section 3).
- [ ] **`yarn build:mac`** artifact tested on target CPU arch; PerfectCue logs show **`(waveshare)`** on connect.
