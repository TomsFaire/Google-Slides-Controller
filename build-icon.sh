#!/bin/bash
# Convert PNG to .icns for macOS app icon
# Usage: ./build-icon.sh path/to/icon.png

if [ -z "$1" ]; then
    echo "Usage: ./build-icon.sh path/to/icon.png"
    echo "The PNG should be at least 1024x1024 pixels for best results"
    exit 1
fi

INPUT="$1"
OUTPUT="build/icon.icns"
ICONSET="build/icon.iconset"

# Check if input file exists
if [ ! -f "$INPUT" ]; then
    echo "Error: File not found: $INPUT"
    exit 1
fi

# Create iconset directory
mkdir -p "$ICONSET"

# Create all required icon sizes for macOS
echo "Creating icon sizes..."

# macOS requires these sizes in the .icns file
sips -z 16 16     "$INPUT" --out "$ICONSET/icon_16x16.png"
sips -z 32 32     "$INPUT" --out "$ICONSET/icon_16x16@2x.png"
sips -z 32 32     "$INPUT" --out "$ICONSET/icon_32x32.png"
sips -z 64 64     "$INPUT" --out "$ICONSET/icon_32x32@2x.png"
sips -z 128 128   "$INPUT" --out "$ICONSET/icon_128x128.png"
sips -z 256 256   "$INPUT" --out "$ICONSET/icon_128x128@2x.png"
sips -z 256 256   "$INPUT" --out "$ICONSET/icon_256x256.png"
sips -z 512 512   "$INPUT" --out "$ICONSET/icon_256x256@2x.png"
sips -z 512 512   "$INPUT" --out "$ICONSET/icon_512x512.png"
sips -z 1024 1024 "$INPUT" --out "$ICONSET/icon_512x512@2x.png"

# Convert iconset to .icns
echo "Converting to .icns..."
iconutil -c icns "$ICONSET" -o "$OUTPUT"

# Clean up iconset directory
rm -rf "$ICONSET"

if [ -f "$OUTPUT" ]; then
    echo "✓ Icon created: $OUTPUT"
    echo "You can now rebuild the app with: npm run build:mac"
else
    echo "✗ Failed to create icon"
    exit 1
fi
