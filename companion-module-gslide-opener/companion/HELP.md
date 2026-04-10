# Companion Module for Google Slides Opener

Control presentations and speaker notes from your Bitfocus Companion setup. Navigate slides, manage speaker notes windows, and monitor real-time status.

## Setup

1. In Companion, import the module from this folder
2. Add a connection and set the presentation computer's IP and API port (default 9595)
3. Create buttons and actions to control your presentations

## What you can do

- Open, close, and reload presentations
- Navigate slides (next, previous, go to specific slide)
- Open and close speaker notes, scroll, and zoom
- Save and recall preset presentations
- Monitor presentation state (open/closed, current slide, etc.)
- **Tunnel QR** – If the Electron app has **Cloudflare Quick Tunnel** enabled (desktop Settings), use **Show Tunnel QR** / **Hide Tunnel QR** to display the tunnel URL as a QR code on the speaker-notes display

## Variables (speaker notes zoom)

The Electron app tracks **native Google Slides** speaker-notes zoom as discrete steps from Slides’ baseline (the same steps as the in-app Zoom in / Zoom out toolbar and the HTTP `/api/zoom-in-notes` / `/api/zoom-out-notes` actions).

- **Speaker Notes Zoom Steps** (`notes_zoom_steps`) — Current offset from baseline while the controller is running (updates when zoom actions succeed and when zoom is restored after reload).
- **Default Speaker Notes Zoom Steps** (`notes_zoom_default`) — The value saved in app preferences (“start larger” when opening notes). Use it to compare the live step count to your configured default.

## Learn more

Visit the main project: [github.com/TomsFaire/Google-Slides-Controller](https://github.com/TomsFaire/Google-Slides-Controller)