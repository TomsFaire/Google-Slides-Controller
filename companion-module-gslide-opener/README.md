# Companion Module for Google Slides Opener

Control the Google Slides Opener app directly from Bitfocus Companion. Navigate slides, manage speaker notes, and monitor presentation state—all from your Companion console.

## What this does

This module connects Bitfocus Companion to Google Slides Opener, giving you full control over presentations and speaker notes. It includes commands for slide navigation, speaker notes management, and preset handling, plus real-time status feedback.

## Installation

Companion 3 modules must be built with **`companion-module-build`** so dependencies (for example `@companion-module/base`) are bundled. Do **not** use `npm pack` (files end up under `package/` and import fails) and do **not** hand-roll a source-only tarball (Companion may import it but the connection will fail with missing modules).

From this folder:

```bash
yarn package
# or: yarn pack:import   (runs the same build, then copies to companion-module-gslide-opener-<version>.tgz)
```

Outputs:

- **`gslide-opener-<version>.tgz`** — primary artifact from the build tool (manifest under `pkg/companion/`, entrypoint in `pkg/`).
- **`companion-module-gslide-opener-<version>.tgz`** — same bytes as above when produced via `yarn pack:import` (handy name for releases/docs).

1. In Bitfocus Companion, go to **Modules** → **Import module package**
2. Select **`gslide-opener-1.4.9.tgz`** or **`companion-module-gslide-opener-1.4.9.tgz`** from the step above (version in the filename matches `package.json`)
3. Add a new connection and configure:
   - **Host**: the IP or hostname of your presentation computer
   - **Port**: the Google Slides Opener API port (default 9595)

That's it. You can now create buttons and actions in Companion to control presentations.

## Available commands

Slide navigation, presentation control, speaker notes, presets, and more. See the main [README](../README.md) for the full list of HTTP API endpoints.

## Learn more

- Main app: [github.com/TomsFaire/Google-Slides-Controller](https://github.com/TomsFaire/Google-Slides-Controller)
- Bitfocus Companion: [bitfocus.io/companion](https://bitfocus.io/companion)