# Installation Instructions for Google Slides Opener

## Important: How to Install on Presentation Machine

The app bundle contains symlinks that must be preserved during installation. Follow these steps carefully:

### Option 1: Use the DMG (Recommended)

1. Copy `Google Slides Opener-1.2.5.dmg` to your presentation machine
2. Double-click the DMG file to mount it
3. Drag `Google Slides Opener.app` from the DMG window to your `/Applications` folder
4. Eject the DMG when done

The DMG preserves all symlinks and structure correctly.

### Option 2: Use the ZIP Archive

1. Copy `Google Slides Opener-1.2.5.zip` to your presentation machine
2. Double-click to extract (macOS Archive Utility preserves symlinks)
3. Move the extracted `Google Slides Opener.app` to `/Applications`

### Option 3: Direct Copy (Use Terminal)

If you must copy directly, use Terminal to preserve symlinks:

```bash
# On your build machine, create a tar archive:
cd /path/to/release
tar -czf Google-Slides-Opener-1.2.5.tar.gz "Google Slides Opener.app"

# Copy the tar.gz to presentation machine, then extract:
cd /Applications
tar -xzf Google-Slides-Opener-1.2.5.tar.gz
```

### Option 4: Using rsync (Preserves Symlinks)

```bash
rsync -av --delete "Google Slides Opener.app" /Applications/
```

## Verification

After installation, verify the app structure is correct:

```bash
# Check if Electron Framework exists
ls -la "/Applications/Google Slides Opener.app/Contents/Frameworks/Electron Framework.framework/Electron Framework"

# Should show a symlink pointing to Versions/Current/Electron Framework
```

If the symlink is broken or missing, the app will crash on launch with:
```
Library not loaded: @rpath/Electron Framework.framework/Electron Framework
```

## Troubleshooting

If the app crashes on launch:

1. **Check symlinks are preserved:**
   ```bash
   ls -la "/Applications/Google Slides Opener.app/Contents/Frameworks/Electron Framework.framework/"
   ```
   You should see `Electron Framework -> Versions/Current/Electron Framework`

2. **Re-sign the app** (if symlinks are broken, this won't fix it, but try):
   ```bash
   codesign --force --deep --sign - "/Applications/Google Slides Opener.app"
   ```

3. **Reinstall using DMG** - This is the most reliable method

## System Requirements

- macOS 10.15 or later
- Apple Silicon (ARM64) or Intel (x64) - Use the appropriate build
- For M4 Macs: Use the ARM64 build
