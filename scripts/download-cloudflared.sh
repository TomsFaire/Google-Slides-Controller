#!/bin/bash
set -euo pipefail
# Pin a version; bump intentionally when upgrading cloudflared.
VERSION="${CLOUDFLARED_VERSION:-2026.3.0}"
mkdir -p resources/cloudflared
BASE="https://github.com/cloudflare/cloudflared/releases/download/${VERSION}"

# macOS: releases ship .tgz containing a binary named "cloudflared"
for pair in "arm64:cloudflared-darwin-arm64" "amd64:cloudflared-darwin-amd64"; do
  arch="${pair%%:*}"
  out="${pair##*:}"
  echo "Downloading ${out} (from .tgz) ..."
  curl -fL -o "resources/cloudflared/${out}.tgz" "${BASE}/cloudflared-darwin-${arch}.tgz"
  tar xzf "resources/cloudflared/${out}.tgz" -C resources/cloudflared
  rm -f "resources/cloudflared/${out}.tgz"
  mv "resources/cloudflared/cloudflared" "resources/cloudflared/${out}"
  chmod +x "resources/cloudflared/${out}"
done

for f in cloudflared-linux-amd64 cloudflared-linux-arm64 cloudflared-windows-amd64.exe; do
  echo "Downloading $f ..."
  curl -fL -o "resources/cloudflared/$f" "${BASE}/$f"
done
chmod +x resources/cloudflared/cloudflared-linux-* 2>/dev/null || true
echo "Done. Version ${VERSION}"
