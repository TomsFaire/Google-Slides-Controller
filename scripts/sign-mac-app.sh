#!/usr/bin/env bash
# Sign the built macOS app so it can be opened without "cannot be opened because of a problem".
# Run from repo root after: npm run build:mac
# Usage: ./scripts/sign-mac-app.sh [mac-arm64|mac-x64]

set -e
ARCH="${1:-mac-arm64}"
APP_PATH="dist/${ARCH}/Google Slides Opener.app"
ENTITLEMENTS="entitlements.mac.plist"

if [ ! -d "$APP_PATH" ]; then
  echo "App not found at $APP_PATH"
  echo "Usage: ./scripts/sign-mac-app.sh [mac-arm64|mac-x64]"
  exit 1
fi

echo "Signing $APP_PATH with hardened runtime + entitlements..."
# Clear quarantine on app root only; -r hits symlinks inside Electron and can trigger ENOTDIR
xattr -c "$APP_PATH" 2>/dev/null || true
find "$APP_PATH" -name "*.framework" -exec codesign --force --sign - --options=runtime {} \;
find "$APP_PATH" -name "*Helper*.app" -exec codesign --force --sign - --options=runtime {} \;
codesign --force --deep --sign - --entitlements "$ENTITLEMENTS" --options=runtime "$APP_PATH"
codesign --verify --verbose "$APP_PATH"
echo "Done. You can open the app now (e.g. open \"$APP_PATH\")."
