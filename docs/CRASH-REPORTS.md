# Crash reports and debug data

The app collects crash data locally so you can attach it when reporting an issue. No data is sent to any server.

## Where reports are stored

- **Crash reports (text):** `userData/crash-reports/`  
  Human-readable `.txt` files for main-process and renderer crashes (error message, stack trace, last 500 log lines).  
  Path is shown in the desktop app under **Debug Logs** → “Crash reports folder”. You can click **Open crash reports folder** to open it.

- **Native crash dumps (minidumps):** Electron’s Crashpad directory, e.g. `userData/crashDumps/` (or the path shown under **Debug Logs** → “Native crash dumps”).  
  Used for native (C++/Electron) crashes. Full stack traces need symbols; these files are still useful to attach for developers.

`userData` is the app’s data directory (e.g. on macOS: `~/Library/Application Support/Google Slides Opener`).

## When a crash happens

- **Main process (app exit):** A dialog explains that a crash report was saved; the app then quits. The report is in `crash-reports/`.
- **Presentation or speaker notes window:** Only that window closes. A dialog may offer to reopen. A “renderer gone” report is written to `crash-reports/`; the rest of the app keeps running.

## Reporting a crash

When opening an issue or contacting support:

1. **Attach** any recent files from the **Crash reports folder** (especially `crash-main-*.txt` or `crash-renderer-*.txt`).
2. **Optionally attach** minidumps from the **Native crash dumps** folder if the app closed with no text report.
3. **Export Debug Log** from the desktop app (Debug Logs → Save Debug Log) and attach that as well if relevant.

This gives developers the error, stack trace, and recent logs needed to diagnose the problem.
