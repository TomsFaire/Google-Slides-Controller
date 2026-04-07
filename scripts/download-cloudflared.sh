#!/bin/bash
set -euo pipefail
# Pin a version; bump intentionally when upgrading cloudflared.
VERSION="${CLOUDFLARED_VERSION:-2026.3.0}"
mkdir -p resources/cloudflared
BASE="https://github.com/cloudflare/cloudflared/releases/download/${VERSION}"
declare -a FILES=(
  cloudflared-darwin-amd64
  cloudflared-darwin-arm64
  cloudflared-linux-amd64
  cloudflared-linux-arm64
  cloudflared-windows-amd64.exe
)
for f in "${FILES[@]}"; do
  echo "Downloading $f ..."
  curl -fL -o "resources/cloudflared/$f" "$BASE/$f"
done
chmod +x resources/cloudflared/cloudflared-darwin-* resources/cloudflared/cloudflared-linux-* 2>/dev/null || true
echo "Done. Version ${VERSION}"
