# Companion Module for Google Slides Opener

Control the Google Slides Opener app directly from Bitfocus Companion. Navigate slides, manage speaker notes, and monitor presentation state—all from your Companion console.

## What this does

This module connects Bitfocus Companion to Google Slides Opener, giving you full control over presentations and speaker notes. It includes commands for slide navigation, speaker notes management, and preset handling, plus real-time status feedback.

## Installation

Build the importable `.tgz` locally (do **not** use `npm pack`—it wraps files in a `package/` folder and Companion will report **missing manifest**):

```bash
yarn pack:import
```

This writes `companion-module-gslide-opener-<version>.tgz` with `companion/manifest.json` at the **root** of the archive.

Alternatively, run **`yarn package`** to produce a webpack-bundled release as `gslide-opener-<version>.tgz` (contents under `pkg/`; Companion accepts this layout).

1. In Bitfocus Companion, go to **Modules** → **Import module package**
2. Select `companion-module-gslide-opener-1.4.9.tgz` (or the version produced above)
3. Add a new connection and configure:
   - **Host**: the IP or hostname of your presentation computer
   - **Port**: the Google Slides Opener API port (default 9595)

That's it. You can now create buttons and actions in Companion to control presentations.

## Available commands

Slide navigation, presentation control, speaker notes, presets, and more. See the main [README](../README.md) for the full list of HTTP API endpoints.

## Learn more

- Main app: [github.com/TomsFaire/Google-Slides-Controller](https://github.com/TomsFaire/Google-Slides-Controller)
- Bitfocus Companion: [bitfocus.io/companion](https://bitfocus.io/companion)