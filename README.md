# Google Slides Opener

A desktop application for quickly switching between Google Slides presentations across multiple monitors, with Bitfocus Companion integration.

## Features

- **Multi-Monitor Support**: Select which monitor to display presentations and notes
- **Google Authentication**: Sign in with your Google account
- **Automated Presentation Mode**: Automatically enters presentation mode and opens speaker notes
- **Quick Window Management**: Press Escape to close both presentation and notes windows
- **HTTP API**: Control presentations remotely via HTTP API (port 9595)
- **Bitfocus Companion Integration**: Control presentations from Stream Deck and other Companion-compatible devices
- **Modern UI**: Clean, intuitive interface

## Installation

1. Install dependencies:
```bash
npm install
```

2. Run the application:
```bash
npm start
```

## Usage

### Using the GUI

1. **Select Monitors**: Choose which monitor for the presentation and which for notes
2. **Sign In**: Click "Sign in with Google" to authenticate
3. **Test**: Use the "Open Test Presentation" button to verify your setup

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

### Bitfocus Companion Integration

This app includes a Companion module for control from Stream Deck and other devices.

See [companion-module-gslide-opener/README.md](companion-module-gslide-opener/README.md) for installation and usage instructions.

**Quick Setup:**
1. Make sure this app is running
2. Copy the `companion-module-gslide-opener` folder to your Companion modules directory
3. Restart Companion
4. Add the "Google Slides Opener" connection in Companion
5. Create buttons with "Open Presentation" actions

## Keyboard Shortcuts

- **Escape**: Close both presentation and notes windows (works from either window)

## Development

```bash
npm run dev
```

## Built With

- Electron
- Node.js HTTP API
- HTML/CSS/JavaScript

## Configuration

Monitor preferences and Google authentication are automatically saved and restored between sessions.
