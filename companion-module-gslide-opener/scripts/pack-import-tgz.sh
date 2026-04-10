#!/usr/bin/env bash
# Companion expects companion/manifest.json at the root of the archive.
# `npm pack` wraps files in package/ which breaks import ("missing manifest").
set -euo pipefail
cd "$(dirname "$0")/.."
NAME=$(node -p "require('./package.json').name")
VER=$(node -p "require('./package.json').version")
OUT="${NAME}-${VER}.tgz"
rm -f "$OUT"
export COPYFILE_DISABLE=1
TAR_FILES=(companion main.js actions.js package.json README.md)
if [[ -f CHANGELOG.md ]]; then TAR_FILES+=(CHANGELOG.md); fi
tar -czvf "$OUT" \
  --exclude='._*' \
  --exclude='.DS_Store' \
  --exclude='**/Icon*' \
  "${TAR_FILES[@]}"
echo "Wrote $(pwd)/${OUT}"
tar -tzf "$OUT"
