#!/usr/bin/env node
/**
 * patch-macadam.js
 *
 * Applies NAPI 10 compatibility patches to macadam C++ source files after
 * yarn install. Required because macadam 2.0.18 source cannot compile against
 * Electron 33 (NAPI 10) without these changes.
 *
 * Patches applied:
 *   capture_promise.cc:
 *     A - finalizeCaptureCarrier: napi_env → node_api_basic_env
 *     B - finalizeVideoBuffer:    napi_env → const napi_env
 *     C - finalizeAudioPacket:    napi_env → const napi_env
 *     D - video buffer creation:  napi_create_external_buffer → napi_create_buffer_copy
 *     E - audio buffer creation:  napi_create_external_buffer → napi_create_buffer_copy
 *   playback_promise.cc:
 *     F - finalizePlaybackCarrier: napi_env → node_api_basic_env
 *
 * Idempotent: re-running after patches are applied is a no-op.
 * Safe when macadam is absent (optionalDependency): exits 0 with a warning.
 */

'use strict';

const fs = require('fs');
const path = require('path');

const MACADAM_SRC = path.join(__dirname, '..', 'node_modules', 'macadam', 'src');
const CAPTURE_FILE = path.join(MACADAM_SRC, 'capture_promise.cc');
const PLAYBACK_FILE = path.join(MACADAM_SRC, 'playback_promise.cc');

// ---------------------------------------------------------------------------
// Patch definitions
// Each patch has:
//   id     - short label for logging
//   file   - absolute path to the file
//   before - exact string that must be present (unpatched form)
//   after  - replacement string (patched form)
// ---------------------------------------------------------------------------
const PATCHES = [
  // --- capture_promise.cc ---------------------------------------------------

  {
    id: 'A',
    file: CAPTURE_FILE,
    before: 'void finalizeCaptureCarrier(napi_env env, void* finalize_data, void* finalize_hint)',
    after:  'void finalizeCaptureCarrier(node_api_basic_env env, void* finalize_data, void* finalize_hint)',
  },

  {
    id: 'B',
    file: CAPTURE_FILE,
    before: 'void finalizeVideoBuffer(napi_env env, void* finalize_data, void* finalize_hint)',
    after:  'void finalizeVideoBuffer(const napi_env env, void* finalize_data, void* finalize_hint)',
  },

  {
    id: 'C',
    file: CAPTURE_FILE,
    before: 'void finalizeAudioPacket(napi_env env, void* finalize_data, void* finalize_hint)',
    after:  'void finalizeAudioPacket(const napi_env env, void* finalize_data, void* finalize_hint)',
  },

  {
    id: 'D',
    // video buffer creation: replace napi_create_external_buffer with buffer-copy pattern
    // Before (single statement):
    //   c->status = napi_create_external_buffer(env, rowBytes*height, bytes, finalizeVideoBuffer, frame->videoFrame, &param);
    // After (copy + immediate release — two lines, 4-space indent):
    //   c->status = napi_create_buffer_copy(env, rowBytes*height, bytes, &bytes, &param);
    //   if (frame->videoFrame) frame->videoFrame->Release();
    file: CAPTURE_FILE,
    before: 'c->status = napi_create_external_buffer(env, rowBytes*height, bytes, finalizeVideoBuffer, frame->videoFrame, &param);',
    // after must exactly match the lines already written in the patched file
    after:  'c->status = napi_create_buffer_copy(env, rowBytes*height, bytes, &bytes, &param);\n    if (frame->videoFrame) frame->videoFrame->Release();',
    // idempotency sentinel: substring that is present only when patch is applied
    applied: 'napi_create_buffer_copy(env, rowBytes*height',
  },

  {
    id: 'E',
    // audio buffer creation: replace napi_create_external_buffer with buffer-copy pattern
    // Before:
    //   c->status = napi_create_external_buffer(env, audioFinalizeData->dataSize, bytes, finalizeAudioPacket, audioFinalizeData, &param);
    // After (three lines, 6-space indent):
    //   c->status = napi_create_buffer_copy(env, audioFinalizeData->dataSize, bytes, &bytes, &param);
    //   if (audioFinalizeData->audioPacket) audioFinalizeData->audioPacket->Release();
    //   free(audioFinalizeData);
    file: CAPTURE_FILE,
    before: 'c->status = napi_create_external_buffer(env, audioFinalizeData->dataSize, bytes, finalizeAudioPacket, audioFinalizeData, &param);',
    after:  'c->status = napi_create_buffer_copy(env, audioFinalizeData->dataSize, bytes, &bytes, &param);\n      if (audioFinalizeData->audioPacket) audioFinalizeData->audioPacket->Release();\n      free(audioFinalizeData);',
    applied: 'napi_create_buffer_copy(env, audioFinalizeData->dataSize',
  },

  // --- playback_promise.cc --------------------------------------------------

  {
    id: 'F',
    file: PLAYBACK_FILE,
    before: 'void finalizePlaybackCarrier(napi_env env, void* finalize_data, void* finalize_hint)',
    after:  'void finalizePlaybackCarrier(node_api_basic_env env, void* finalize_data, void* finalize_hint)',
  },
];

// ---------------------------------------------------------------------------
// Apply patches
// ---------------------------------------------------------------------------

let missingFiles = false;
const fileCache = {};

function readFile(filePath) {
  if (!(filePath in fileCache)) {
    fileCache[filePath] = fs.readFileSync(filePath, 'utf8');
  }
  return fileCache[filePath];
}

function writeFile(filePath, content) {
  fileCache[filePath] = content;
  fs.writeFileSync(filePath, content, 'utf8');
}

// Check required files exist first
for (const p of [CAPTURE_FILE, PLAYBACK_FILE]) {
  if (!fs.existsSync(p)) {
    missingFiles = true;
  }
}

if (!fs.existsSync(MACADAM_SRC) || missingFiles) {
  console.warn('[patch-macadam] macadam source not found — skipping patches (optional dependency absent).');
  process.exit(0);
}

let allOk = true;

for (const patch of PATCHES) {
  const shortFile = path.relative(path.join(__dirname, '..'), patch.file);
  const content = readFile(patch.file);

  // Use explicit applied sentinel if provided, otherwise fall back to after-text
  const alreadyApplied = patch.applied
    ? content.includes(patch.applied)
    : content.includes(patch.after);

  if (alreadyApplied) {
    console.log(`[patch-macadam] Patch ${patch.id} already applied in ${shortFile} — skipped.`);
    continue;
  }

  if (!content.includes(patch.before)) {
    // Neither before nor after found — unexpected state
    console.error(
      `[patch-macadam] ERROR: Patch ${patch.id}: expected string not found in ${shortFile}.\n` +
      `  Looking for: ${patch.before.slice(0, 80)}...`
    );
    allOk = false;
    continue;
  }

  const patched = content.replace(patch.before, patch.after);
  writeFile(patch.file, patched);
  console.log(`[patch-macadam] Patch ${patch.id} applied to ${shortFile}.`);
}

if (!allOk) {
  console.error('[patch-macadam] One or more patches failed. The macadam native build may fail.');
  process.exit(1);
}

console.log('[patch-macadam] All patches applied successfully.');
