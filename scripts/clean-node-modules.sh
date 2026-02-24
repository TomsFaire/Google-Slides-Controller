#!/usr/bin/env bash
# Remove node_modules when "rm -rf" fails (nested .bin symlinks etc. on macOS).
# Order: symlinks first, then files, then empty dirs.
set -e
ROOT="${1:-.}"
NM="${ROOT}/node_modules"
if [ ! -d "$NM" ]; then
  echo "No node_modules at $NM"
  exit 0
fi
echo "Removing $NM (symlinks, then files, then empty dirs)..."
find "$NM" -type l -delete 2>/dev/null || true
find "$NM" -type f -delete 2>/dev/null || true
find "$NM" -depth -type d -empty -exec rmdir {} \; 2>/dev/null || true
# If anything is left (e.g. non-empty dirs), do a depth-first delete
if [ -d "$NM" ] && [ -n "$(ls -A "$NM" 2>/dev/null)" ]; then
  find "$NM" -depth -mindepth 1 -delete 2>/dev/null || true
fi
echo "Done."
