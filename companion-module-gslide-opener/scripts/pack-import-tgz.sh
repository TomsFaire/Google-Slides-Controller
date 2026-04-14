#!/usr/bin/env bash
# Companion 3 expects a package from `companion-module-build` (webpack bundle + pkg/ layout).
# A plain `tar` of source files has no bundled deps — import may succeed but the module will
# fail to load (`Cannot find module '@companion-module/base'`).
set -euo pipefail
cd "$(dirname "$0")/.."
NAME=$(node -p "require('./package.json').name")
VER=$(node -p "require('./package.json').version")
BUILT="gslide-opener-${VER}.tgz"
OUT="${NAME}-${VER}.tgz"

echo "[pack-import] Running companion-module-build (yarn package)..."
yarn run package

if [[ ! -f "$BUILT" ]]; then
  echo "ERROR: expected $BUILT after build" >&2
  exit 1
fi

cp -f "$BUILT" "$OUT"
echo "[pack-import] Wrote $(pwd)/$OUT"
echo "[pack-import] (same contents as $BUILT; rename for release/docs compatibility)"
tar -tzf "$OUT" | head -20
