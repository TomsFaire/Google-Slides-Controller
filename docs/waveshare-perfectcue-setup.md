# WaveShare serial–Ethernet + PerfectCue (operator checklist)

Use this with the **WaveShare** branch settings: each PerfectCue listener row can be **DSAN** or **WaveShare** so keep-alive timers match the converter.

## Converter network settings

- **Work mode:** TCP **Client** (connect out to the machine running Google Slides Controller).
- **Destination IP:** The controller host’s LAN IP (not a typo—must match the PC/Mac running the app).
- **Destination port:** The **same** TCP port shown under PerfectCue listener ports in app settings (default **8899**). This is **not** the HTTP API port (9595).

## Serial / RS485

- The WaveShare **DB9 is RS-232**—use the **screw-terminal block** for **RS485** (A/B/GND per silkscreen).
- Match **baud and framing** to the PerfectCue output on that link (often **115200 8N1** on WaveShare deployments; DSAN USR modules may differ—confirm with hardware).
- Wire **RS485** per WaveShare and PerfectCue documentation (A/B, termination if required).

## Field validation (ties to release QA)

1. **Destination IP/port** — Confirm the converter reaches the listener (`[PerfectCue] Listening on port …` and `client connected (waveshare)` in debug logs).
2. **Baud / framing** — Confirm `[PerfectCue] raw:` lines are plausible hex (not garbage); adjust baud if needed.
3. **Command bytes** — Confirm cues emit **`0x0F` / `0x1F`** per byte (current parser). If you only see ASCII lines, a parser update is required—capture logs for developers.
4. **Mixed adapters** — Run two listener ports (one **DSAN**, one **WaveShare**) and idle-soak both; neither tunnel should drop idle bridges nor require hardware reboot when only the app restarts.

Presets live in `src/perfectcue-adapter-presets.js` (`dsan` vs `waveshare` ping and read-idle timeouts).
