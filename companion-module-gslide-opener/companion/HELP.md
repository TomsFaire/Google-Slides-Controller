# Google Slides Opener

Control Google Slides presentations with automatic presentation mode and multi-monitor support.

## Setup

1. Make sure the Google Slides Opener Electron app is running
2. The app will start an HTTP API server on port 9595
3. Configure the connection in Companion with host `127.0.0.1` and port `9595`

## Actions

### Open Presentation
Opens a Google Slides presentation and automatically:
- Closes any previously open presentation
- Enters presentation mode
- Opens speaker notes
- Positions windows on configured monitors

**Options:**
- **Google Slides URL**: The full URL to your presentation (supports variables)

### Close Current Presentation
Closes the currently open presentation and speaker notes windows.

### Next Slide
Advances to the next slide in the presentation.

### Previous Slide
Goes back to the previous slide in the presentation.

### Toggle Video Playback
Toggles video playback on the current slide (if a video is embedded). Uses the 'k' hotkey.

### Zoom In Speaker Notes
Increases the zoom level of the speaker notes window.

### Zoom Out Speaker Notes
Decreases the zoom level of the speaker notes window.

## Configuration

Monitor preferences are configured in the Google Slides Opener app GUI. The Companion module will use those saved preferences when opening presentations.

## Tips

- Use Companion variables in the URL field for dynamic presentation switching
- Create multiple buttons with different presentation URLs
- The app must be signed in to Google for presentations to open
- Press Escape in either window to close both presentation and notes

## Troubleshooting

If the connection fails:
1. Check that the Google Slides Opener app is running
2. Verify the app shows: `[API] HTTP server listening on http://127.0.0.1:9595`
3. Check that port 9595 is not blocked by a firewall
4. Test the API directly: open `http://127.0.0.1:9595/api/status` in your browser
