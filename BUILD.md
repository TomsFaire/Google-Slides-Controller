# Building Google Slides Opener

This project uses GitHub Actions to automatically build portable executables for Windows and Linux.

## Automated Builds

### When builds run:
- On every push to `main` or `master` branch
- On pull requests
- On manual trigger (via GitHub Actions tab)
- On version tags (e.g., `v1.0.0`) - automatically creates a GitHub Release

### Build artifacts:
- **Windows**: Portable `.exe` (no installer needed)
- **Linux**: `.AppImage` (portable, no installation needed)
- **Companion Module**: `.tgz` package for Bitfocus Companion integration

## Manual Building

### Prerequisites
```bash
npm install
```

### Build commands:

**On Windows:**
```bash
# Build Windows portable exe
npm run build:win
```

**On Linux:**
```bash
# Build Linux AppImage
npm run build:linux
```

**Cross-platform builds:** You cannot build Linux binaries on Windows or vice versa. Use GitHub Actions to build for both platforms automatically.

Built files will be in the `dist/` directory.

**Note for Windows:** The build process disables code signing automatically. The first build may take 5-10 minutes while dependencies are downloaded and cached. Subsequent builds will be much faster.

## Optional: App Icons

To customize the app icon, create a `build/` directory and add:
- `build/icon.ico` - Windows icon (256x256 recommended)
- `build/icon.png` - Linux icon (512x512 recommended)

If no icons are provided, electron-builder will use default icons.

## Creating a Release

To create a GitHub Release with downloadable builds:

1. Create and push a version tag:
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```

2. GitHub Actions will automatically:
   - Build Windows portable exe
   - Build Linux AppImage
   - Package Companion module as .tgz
   - Create a GitHub Release
   - Attach all builds to the release

## Downloading Builds

### From GitHub Actions:
1. Go to the "Actions" tab in your repository
2. Click on the latest successful workflow run
3. Scroll down to "Artifacts" section
4. Download:
   - `gslide-opener-windows-portable` for Windows
   - `gslide-opener-linux-appimage` for Linux
   - `companion-module-gslide-opener` for Bitfocus Companion

### From GitHub Releases:
1. Go to the "Releases" section
2. Download the latest release assets

## Troubleshooting

### Build fails on Linux
Make sure you have the required dependencies:
```bash
sudo apt-get install -y libarchive-tools
```

### Build fails on Windows
- Ensure you have the latest npm and Node.js installed
- The first build may take 5-10 minutes - this is normal
- If you get "The process cannot access the file", close any running instances of the app and delete the `dist` folder before rebuilding

### Cannot build Linux on Windows (or vice versa)
This is expected - you need the target OS to build binaries for that OS:
- **Windows builds** require Windows
- **Linux builds** require Linux
- **Solution:** Use GitHub Actions which automatically builds for both platforms
