# DeckLink output

Slides and speaker notes can be sent out of a Blackmagic DeckLink / UltraStudio device as a video signal. This page covers setup, how it works, and troubleshooting. The short version is in the [README](../README.md#decklink-output).

## Requirements

- A DeckLink or UltraStudio device with video output.
- Blackmagic Desktop Video installed and the device visible in **Blackmagic Desktop Video Setup**. Verified with Desktop Video 16.4.
- The device is not in use by another application.
- macOS 13+ (release builds). Windows and Linux are untested.

The app talks to the device through [`macadam`](https://www.npmjs.com/package/macadam), a native module that is bundled with the release builds. If it cannot load, the DeckLink tab shows why and no output is available. There is no fallback. The Provider line in the tab shows the installed DeckLink API version, which tracks the Desktop Video release.

## Using it

1. Desktop settings → **DeckLink** tab.
2. For slides and/or notes: tick **Enable**, choose the **Device** and **Format**.
3. **Save DeckLink Settings**. The status light turns on when the output is live.

Pick a format that matches the receiving equipment. Many receivers will not lock to a rate they do not expect.

The Format list is read from the selected device, so it only shows modes that device supports. The app offers up to these: 1080p 59.94 / 60 / 50 / 30 / 29.97 / 25, 720p 59.94 / 50 (UHD and higher are not supported yet and are hidden even if the device has them). If the device cannot be read, the full list is shown. If a saved format is not supported by the device you pick, the tab switches to a supported one and says so.

The app asks the device which pixel format it accepts for the chosen mode. It sends BGRA when the device takes it, and otherwise converts to 8-bit YUV itself (some devices, such as the UltraStudio Express Monitor 3G, accept 1080p50/59.94/60 only as YUV). The log line `[decklink] output started: … pixel format …` shows which was used.

## What it does not do

- **No audio.** The output is video only. Route audio from the computer into your mixer.
- **No UHD / 4K or 10-bit output.**

## Performance

The app captures the window with Electron's `capturePage()`, resizes to the output size if needed, and pushes the frame to a worker process that owns the DeckLink device. A crash in the DeckLink driver kills the worker, not the app.

Every 5 seconds the app logs, per output:

```
[decklink] slides: 21.2 fps captured, capture 12.7ms, resize 33.8ms, 0 ticks skipped (capture busy), 53 frames dropped (device busy), worker ...
```

- **capture / resize:** time per frame. Resize is 0 ms when the window's pixels already match the output size, for example a 1080p display at 1× scale. On a 4K display (a 3840×2160 capture downscaled to 1080p) we measured about 12.7 ms capture and about 34 ms resize, which capped capture at about 21 fps. Resize runs on the main thread, so it limits the frame rate no matter how fast the device is.
- **frames dropped (device busy):** the worker was still sending the previous frame.
- **ticks skipped (capture busy):** the previous capture had not finished.

Static slides are smooth. For video and animation, check the log with your real content before an event. The target is 24 fps minimum, 30 where possible; this has not yet been measured with video on a 1080p display.

### Measured baseline (macadam backend)

UltraStudio Express Monitor 3G, Desktop Video 16.4.0, 1080p29.97 (BGRA), 2026-10-23. Recorded so the native addon can be compared against it.

| Source | Captured | Capture | Resize | Delivered by worker | Round trip | `displayFrame` |
|---|---|---|---|---|---|---|
| Fake window (colour bars, no capture or resize) | 29.8 fps | 0 ms | 0 ms | 29.4 fps | 29.1 ms | 21.1 ms |
| Real slides on a UHD display | 16 to 19 fps | 13 to 23 ms | about 37 ms | 8 to 9.6 fps | 53 to 61 ms | about 21 ms |

On the UHD display the main thread is the limit: capture plus resize is 50 to 60 ms per frame. Next-slide changes and animated GIFs looked fine on the output. Video slides on a 1080p 1x display are still unmeasured.

## Colour

Frames are converted to YUV using Rec. 709, limited range, which is right for HD monitors and mixers. If a display looks different from a computer screen, check the display first: run `yarn decklink:color-chart` and compare against SMPTE-style bars on a waveform / vectorscope if you have one. Environment variables `DECKLINK_YUV_MATRIX` (`709`|`601`) and `DECKLINK_YUV_RANGE` (`limited`|`full`) exist for A/B testing only.

## Troubleshooting

| Symptom | Check |
|---|---|
| Provider status says unavailable | Desktop Video installed? Device visible in Desktop Video Setup? Restart the app. |
| "Unable to access the hardware" | Another application holds the device. Quit it. If it was the app, quit and relaunch. |
| "Requested display mode is not supported" | The device does not support that format. Try another. `yarn decklink:probe-modes` lists what works. |
| Signal comes up in a different mode | Desktop Video Setup can override the output format. Check it there. |
| Choppy video / animation | See Performance. Try a 1080p display at 1× scale, or a lower format. |
| Crash on first use | Desktop Video 16+ needs the app's patched macadam build. Use a release build. |

## Diagnostic tools

Run from a source checkout. These need the DeckLink device, and they use Electron's Node, so run them through `yarn`.

| Command | Purpose |
|---|---|
| `yarn decklink:info` | Read-only dump of the device's capabilities and supported modes. |
| `yarn decklink:probe-modes` | Tries each format the app offers and reports which the device accepts. |
| `yarn test:decklink-hw` | Hardware test (colour bars, start/stop cycles). Set `DECKLINK_MODE`, `DECKLINK_DEVICE`, `DECKLINK_HOLD_MS`. |
| `yarn decklink:color-chart` | Outputs bars and a greyscale for checking colour on a monitor. |
| `yarn decklink:capture-colors` | Compares window capture colours per display. |

## Building from source

`yarn install` runs `scripts/patch-macadam.js`, which patches macadam so it builds against current Electron and Desktop Video. If `yarn install` fails building macadam with your system Node, use:

```bash
yarn install --ignore-scripts && node scripts/patch-macadam.js && yarn rebuild
```
