# Google Slides Opener

A desktop application for quickly switching between Google Slides presentations across multiple monitors, with Bitfocus Companion integration.

![Demo](docs/images/demo.gif)

## Installation

1. Download the latest release of the Google Slides Opener (exe or Appimage) on your presentation machine, and download the `companion-module-gslide-opener.tgz` on your Companion machine.
2. Open the Google Slides Opener executable and make any config changes you want.
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
- `POST /api/toggle-video` - Toggle video playback
- `POST /api/open-speaker-notes` - Open or close speaker notes (s key)
- `POST /api/close-speaker-notes` - Close the speaker notes window
- `POST /api/zoom-in-notes` - Zoom in on speaker notes
- `POST /api/zoom-out-notes` - Zoom out on speaker notes

#### Example with curl

```bash
# Open a presentation
curl -X POST http://127.0.0.1:9595/api/open-presentation \
  -H "Content-Type: application/json" \
  -d '{"url":"https://docs.google.com/presentation/d/YOUR_ID/edit"}'

# Close presentation
curl -X POST http://127.0.0.1:9595/api/close-presentation
```