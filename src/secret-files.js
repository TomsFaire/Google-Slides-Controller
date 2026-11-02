const fs = require('fs');

/** Owner read/write only. Everything holding a secret on disk gets this. */
const SECRET_FILE_MODE = 0o600;

/**
 * True when a mode grants any group or other permission.
 * @param {number} mode st_mode from fs.statSync
 * @returns {boolean}
 */
function isModeTooPermissive(mode) {
  return (mode & 0o077) !== 0;
}

/**
 * Write a file holding a secret, then tighten it to owner-only.
 * The `mode` option of writeFileSync is only applied when the file is created,
 * so a file that already exists keeps whatever mode it had; the explicit chmod
 * is what fixes the files written by earlier versions.
 * @param {string} filePath
 * @param {string|Buffer} data
 * @param {string} [encoding] passed through to writeFileSync when given
 */
function writeSecretFileSync(filePath, data, encoding) {
  const options = encoding ? { encoding, mode: SECRET_FILE_MODE } : { mode: SECRET_FILE_MODE };
  fs.writeFileSync(filePath, data, options);
  try {
    fs.chmodSync(filePath, SECRET_FILE_MODE);
  } catch (_) {
    // Best effort: a filesystem that cannot represent the mode (e.g. a Windows
    // volume or a FAT-formatted drive) should not fail the write itself.
  }
}

/**
 * Tighten an existing secret file written by an earlier version.
 * @param {string} filePath
 * @returns {'missing'|'already-owner-only'|'tightened'|'failed'}
 */
function ensureSecretFileMode(filePath) {
  try {
    if (!fs.existsSync(filePath)) return 'missing';
    const { mode } = fs.statSync(filePath);
    if (!isModeTooPermissive(mode)) return 'already-owner-only';
    fs.chmodSync(filePath, SECRET_FILE_MODE);
    return 'tightened';
  } catch (_) {
    return 'failed';
  }
}

module.exports = {
  SECRET_FILE_MODE,
  isModeTooPermissive,
  writeSecretFileSync,
  ensureSecretFileMode
};
