# Deployment Guide - Google Slides Opener

## What You Need to Download

**The GitHub repository only contains SOURCE CODE.** The built application files are NOT in GitHub (they're in `.gitignore`).

To run the app on your presentation machine, you need ONE of these files from the `release/` folder:

### Recommended: DMG File
- **File:** `Google Slides Opener-1.2.5.dmg` (~100MB)
- **Location:** `/Users/tom/Documents/gslide-opener/release/`
- **Why:** Best preserves symlinks and app structure
- **How to use:**
  1. Copy the DMG file to your presentation machine
  2. Double-click to mount it
  3. Drag `Google Slides Opener.app` to `/Applications`
  4. Eject the DMG

### Alternative: ZIP Archive
- **File:** `Google Slides Opener-1.2.5.zip` (~267MB)
- **Location:** `/Users/tom/Documents/gslide-opener/release/`
- **How to use:**
  1. Copy ZIP file to presentation machine
  2. Double-click to extract
  3. Move `Google Slides Opener.app` to `/Applications`

### Alternative: TAR.GZ Archive
- **File:** `Google-Slides-Opener-1.2.5.tar.gz` (~90MB)
- **Location:** `/Users/tom/Documents/gslide-opener/release/`
- **How to use:**
  ```bash
  cd /Applications
  tar -xzf ~/Downloads/Google-Slides-Opener-1.2.5.tar.gz
  ```

### Direct: App Bundle (Advanced)
- **File:** `Google Slides Opener.app` folder (~224MB)
- **Location:** `/Users/tom/Documents/gslide-opener/release/`
- **Warning:** Must preserve symlinks when copying!
- **How to use:**
  ```bash
  # Use rsync to preserve symlinks:
  rsync -av "Google Slides Opener.app" /Applications/
  ```

## Where to Find These Files

On your **build machine** (where you developed the app):
```
/Users/tom/Documents/gslide-opener/release/
├── Google Slides Opener-1.2.5.dmg          ← RECOMMENDED
├── Google Slides Opener-1.2.5.zip          ← Alternative
├── Google-Slides-Opener-1.2.5.tar.gz      ← Alternative
└── Google Slides Opener.app/               ← Direct (advanced)
```

## What NOT to Download

- ❌ Don't download from GitHub - it only has source code
- ❌ Don't copy just the MacOS binary - you need the entire .app bundle
- ❌ Don't use Finder drag-and-drop - it may break symlinks

## Quick Start

1. **On your build machine:** Copy `Google Slides Opener-1.2.5.dmg` to a USB drive or network share
2. **On presentation machine:** Copy the DMG file to Downloads
3. **Double-click the DMG** to mount it
4. **Drag the app** to Applications folder
5. **Launch** the app from Applications

## Companion Module

The Companion module is in a separate file:
- **File:** `gslide-opener-1.2.5.tgz`
- **Location:** `/Users/tom/Documents/gslide-opener/release/`
- **Install in Companion:** Use Companion's module installer with this .tgz file
