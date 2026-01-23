#!/bin/bash
# Test script to verify code signing works for downloaded apps
# This simulates what happens when a user downloads the app from GitHub

set -e

echo "=== Testing Code Signing Process ==="
echo ""

# Step 1: Build the app
echo "Step 1: Building the app..."
npm run build:mac

# Step 2: Find the built app
APP_PATH="dist/mac-arm64/Google Slides Opener.app"
ZIP_PATH=$(ls -1 dist/*.zip | head -1)

if [ ! -d "$APP_PATH" ]; then
    echo "ERROR: App not found at $APP_PATH"
    exit 1
fi

echo "✓ App built at: $APP_PATH"
echo "✓ ZIP created at: $ZIP_PATH"
echo ""

# Step 3: Simulate download by adding quarantine attribute
echo "Step 2: Simulating download (adding quarantine attribute)..."
xattr -w com.apple.quarantine "0081;$(date +%s);;$(uuidgen)" "$APP_PATH"
echo "✓ Quarantine attribute added"
echo ""

# Step 4: Try to verify signature (should fail or show issues)
echo "Step 3: Checking signature before signing..."
codesign --verify --verbose "$APP_PATH" 2>&1 || echo "⚠ Signature check failed (expected if not signed)"
echo ""

# Step 5: Apply our signing process
echo "Step 4: Applying signing process..."
echo "  - Removing quarantine..."
xattr -cr "$APP_PATH"
echo "  - Signing nested frameworks and helpers..."
find "$APP_PATH" -name "*.framework" -exec codesign --force --sign - {} \;
find "$APP_PATH" -name "*Helper*.app" -exec codesign --force --sign - {} \;
echo "  - Ad-hoc signing main app..."
codesign --force --deep --sign - "$APP_PATH"
echo "  - Verifying signature..."
codesign --verify --verbose "$APP_PATH"
echo "✓ Signing complete"
echo ""

# Step 6: Check signature details
echo "Step 5: Signature details:"
codesign -dv "$APP_PATH" 2>&1
echo ""

# Step 7: Test the ZIP recreation process
echo "Step 6: Testing ZIP recreation with signed app..."
TEMP_DIR=$(mktemp -d)
cd dist

# Extract existing ZIP
ZIP_FILE=$(basename "$ZIP_PATH")
unzip -q "$ZIP_FILE" -d "$TEMP_DIR"

# The app in the ZIP should be unsigned (from electron-builder)
echo "  - Checking app in original ZIP..."
APP_IN_ZIP="$TEMP_DIR/Google Slides Opener.app"
if [ -d "$APP_IN_ZIP" ]; then
    codesign --verify --verbose "$APP_IN_ZIP" 2>&1 || echo "  ⚠ App in ZIP is not signed (expected)"
fi

# Remove old ZIP
rm -f "$ZIP_FILE"

# Re-zip with signed app
cd mac-arm64
ditto -c -k --keepParent "Google Slides Opener.app" "../$ZIP_FILE"
cd ..

# Verify the app in the new ZIP
echo "  - Extracting and verifying signed app in new ZIP..."
rm -rf "$TEMP_DIR"
TEMP_DIR=$(mktemp -d)
unzip -q "$ZIP_FILE" -d "$TEMP_DIR"
NEW_APP="$TEMP_DIR/Google Slides Opener.app"

if [ -d "$NEW_APP" ]; then
    # Add quarantine to simulate download
    xattr -w com.apple.quarantine "0081;$(date +%s);;$(uuidgen)" "$NEW_APP"
    echo "  - Added quarantine (simulating download)..."
    
    # Remove quarantine and verify
    xattr -cr "$NEW_APP"
    codesign --verify --verbose "$NEW_APP" && echo "  ✓ App in ZIP is properly signed!"
fi

rm -rf "$TEMP_DIR"
cd ..
echo ""

echo "=== Test Complete ==="
echo ""
echo "To test if the app actually runs:"
echo "  1. Extract the ZIP: unzip -q '$ZIP_PATH' -d /tmp/test-app"
echo "  2. Try to open it: open '/tmp/test-app/Google Slides Opener.app'"
echo "  3. If you see 'damaged' error, apply: xattr -cr '/tmp/test-app/Google Slides Opener.app'"
echo "  4. Then try opening again"
echo ""
echo "The signing process should prevent the 'damaged' error."
