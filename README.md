# Google Slides Opener

A desktop application for quickly switching between Google Slides presentations across multiple monitors, with Bitfocus Companion integration.

![Demo](docs/images/demo.gif)

## Attribution

This project is based on [nerif-tafu/gslide-opener](https://github.com/nerif-tafu/gslide-opener).

Original work by [nerif-tafu](https://github.com/nerif-tafu). This fork includes additional features and improvements.

## Installation

1. Download the latest release of the Google Slides Opener (exe or Appimage) on your presentation machine, and download the `companion-module-gslide-opener.tgz` on your Companion machine.
2. **macOS users:** On first launch, you may see a security warning. This is expected behavior. Right-click the app and select "Open" to launch it. See [README-SECURITY.md](README-SECURITY.md) for details.
3. Open the Google Slides Opener executable and make any config changes you want:
   - Select your presentation and notes monitors
   - Configure network ports (API and Web UI)
   - View available network IP addresses for accessing the Web UI
3. On your Companion instance, go to `Modules` > `Import module package` and select the .tgz file.
4. Add the connection in Companion named `Google Slides Opener`

## Development

To work on this repository run the following commands:

```bash
npm install
npm run dev
```

If you would like to test the build you can use the following commands: 

```bash
npm run build:win # Builds the .exe for windows
npm run build:linux # Builds the appimage for Linux
./package-companion.ps1 # Builds the companion .tgz
```

### Web UI for Preset Management and Controls

The app includes a web interface (default port **80**, configurable) for managing preset presentations and controlling active presentations:

- **Access:** Open `http://YOUR_COMPUTER_IP` (or `http://YOUR_COMPUTER_IP:PORT`) in any web browser
- **Port Configuration:** The web UI port can be configured in the desktop app's "Network Ports" section
- **Features:**

  **Controls Tab:**
  - **Slide Controls:**
    - Previous Slide
    - Next Slide
    - Reload Presentation (preserves current slide and zoom level)
  - **Speaker Notes Controls:**
    - Zoom In / Zoom Out
    - Scroll Up / Scroll Down

  **Presets Tab:**
  - Configure 3 preset presentations (Presentation 1, 2, 3)
  - Save and load presets
  - These presets can then be opened from Companion using the "Open Presentation 1/2/3" actions without needing to enter URLs

- **Network Access:** The desktop app displays all available IP addresses where the web UI can be accessed in the "Network Access" section
- **Accessible from any device:** Use the web UI from phones, tablets, or other computers on your network

### Using the HTTP API

The app exposes an HTTP API on `http://127.0.0.1:9595` when running:

#### Endpoints

- `GET /api/status` - Check if the app is running
- `POST /api/open-presentation` - Open a presentation
  ```json
  {
    "url": "https://docs.google.com/presentation/d/YOUR_ID/edit"
  }
  ```
- `POST /api/close-presentation` - Close current presentation
- `POST /api/next-slide` - Go to next slide
- `POST /api/previous-slide` - Go to previous slide
- `POST /api/go-to-slide` - Navigate to a specific slide number
  ```json
  {
    "slide": 5
  }
  ```
- `POST /api/reload-presentation` - Close and reopen the current presentation, returning to the same slide
- `POST /api/toggle-video` - Toggle video playback
- `POST /api/open-speaker-notes` - Open or close speaker notes (s key)
- `POST /api/close-speaker-notes` - Close the speaker notes window
- `POST /api/scroll-notes-down` - Scroll speaker notes down (150px)
- `POST /api/scroll-notes-up` - Scroll speaker notes up (150px)
- `POST /api/zoom-in-notes` - Zoom in on speaker notes
- `POST /api/zoom-out-notes` - Zoom out on speaker notes
- `GET /api/presets` - Get all preset presentation URLs
- `POST /api/presets` - Set preset presentation URLs
  ```json
  {
    "presentation1": "https://docs.google.com/presentation/d/...",
    "presentation2": "https://docs.google.com/presentation/d/...",
    "presentation3": "https://docs.google.com/presentation/d/..."
  }
  ```
- `POST /api/open-preset` - Open a preset by number (1, 2, or 3)
  ```json
  {
    "preset": 1
  }
  ```

#### Example with curl

```bash
# Open a presentation
curl -X POST http://127.0.0.1:9595/api/open-presentation \
  -H "Content-Type: application/json" \
  -d '{"url":"https://docs.google.com/presentation/d/YOUR_ID/edit"}'

# Close presentation
curl -X POST http://127.0.0.1:9595/api/close-presentation
```
