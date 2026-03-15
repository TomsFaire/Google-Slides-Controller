# Installing Google Slides Opener

The app has internal symlinks that need to be preserved during installation. Here's how to do it right.

## The easy way: DMG (recommended)

1. Copy the `.dmg` file to your presentation computer
2. Double-click it to mount
3. Drag the app to `/Applications`
4. Eject when done

That's it. The DMG format preserves all the internal structure.

## The quick way: ZIP

1. Copy the `.zip` to your presentation computer
2. Double-click to extract
3. Move the app to `/Applications`

macOS's built-in Archive Utility handles symlinks correctly.

## The terminal way: tar

If you're copying over the network:

```bash
# Create a tar archive on the source machine
cd /path/to/release
tar -czf Google-Slides-Opener.tar.gz "Google Slides Opener.app"

# Extract on the destination machine
cd /Applications
tar -xzf Google-Slides-Opener.tar.gz
```

## Or use rsync

```bash
rsync -av "Google Slides Opener.app" /Applications/
```

## Verify it worked

After installation, check that the internal structure is intact:

```bash
ls -la "/Applications/Google Slides Opener.app/Contents/Frameworks/Electron Framework.framework/Electron Framework"
```

You should see `Electron Framework -> Versions/Current/Electron Framework` (a symlink, indicated by the arrow).

If the symlink is missing or broken, you'll get this error on launch:
```
Library not loaded: @rpath/Electron Framework.framework/Electron Framework
```

## App won't launch?

1. **First, check the symlink:**
   ```bash
   ls -la "/Applications/Google Slides Opener.app/Contents/Frameworks/Electron Framework.framework/"
   ```
   If you don't see the symlink, reinstall using the DMG.

2. **Try re-signing** (sometimes helps):
   ```bash
   codesign --force --deep --sign - "/Applications/Google Slides Opener.app"
   ```

3. **Reinstall with DMG** — this is the most reliable fix.

## System Requirements

- macOS 10.15+
- ARM64 (Apple Silicon) or x64 (Intel) — choose the matching build
- M1/M2/M3/M4 Macs use the ARM64 version
