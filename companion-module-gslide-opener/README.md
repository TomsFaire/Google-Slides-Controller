# Bitfocus Companion Module for Google Slides Opener

This module allows you to control the Google Slides Opener Electron app from Bitfocus Companion.

## Features

- **Open Presentation**: Open a Google Slides presentation with a custom URL
- **Close Presentation**: Close the currently open presentation
- Automatically closes previous presentations when opening a new one
- Uses saved monitor preferences from the Electron app

## Setup

### 1. Install the Electron App

Make sure the Google Slides Opener Electron app is running on your computer. The app starts an HTTP API server on port 9595.

### 2. Install the Companion Module

1. In Companion web interface, go to **Settings**
2. Set **Developer modules path** to the parent folder containing `companion-module-gslide-opener`
   - Example: `C:\Users\YourName\Work Repos\gslide-opener`
3. Install module dependencies:
   ```bash
   cd companion-module-gslide-opener
   npm install
   ```
4. Restart Bitfocus Companion

### 3. Add the Module to Companion

1. Open Bitfocus Companion web interface
2. Go to the "Connections" tab
3. Click "Add Connection"
4. Search for "Google Slides Opener"
5. Configure the connection:
   - **Host**: `127.0.0.1` (default)
   - **Port**: `9595` (default)
6. The status should show as "OK" if the Electron app is running

## Usage

### Setting up a Button

1. Create a new button in Companion
2. Add an action: "Open Presentation"
3. Enter the Google Slides URL (e.g., `https://docs.google.com/presentation/d/YOUR_PRESENTATION_ID/edit`)
4. The button will now open that presentation when pressed

### Using Variables

You can use Companion variables in the URL field. For example:
- Create a custom variable with your presentation URL
- Use `$(custom:my_presentation_url)` in the URL field

### Multiple Presentations

Set up multiple buttons with different URLs to quickly switch between presentations. Each button press will:
1. Close any currently open presentation
2. Open the new presentation on your configured monitors
3. Automatically enter presentation mode
4. Open speaker notes

## Troubleshooting

- **"Connection Failed"**: Make sure the Google Slides Opener app is running
- **"Timeout"**: Check that the port (9595) is not blocked by a firewall
- **Presentation doesn't open**: Check the Companion logs for error messages

## API Endpoints

The module communicates with these endpoints:

- `GET /api/status` - Check if the app is running
- `POST /api/open-presentation` - Open a presentation with a URL
- `POST /api/close-presentation` - Close the current presentation

## Development

To modify this module:

1. Edit the files in this directory
2. Restart Companion to reload changes
3. Check the Companion logs for debugging information

## Support

For issues or questions, please check the main project repository.
