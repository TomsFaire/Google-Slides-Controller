/**
 * How well safeStorage is actually protecting secrets on this machine.
 *
 * Electron reports two things that matter and that do not always agree:
 *   - isEncryptionAvailable(), which says whether safeStorage will encrypt at all;
 *   - getSelectedStorageBackend(), Linux-only, which names the secret store.
 *
 * The case worth warning about is 'basic_text': Electron still returns a
 * base64 blob that looks encrypted, but the key is a hardcoded password, so it
 * is not meaningful protection. That shows up on Linux boxes with no keyring
 * (a headless AV machine, typically), where it is the normal result rather
 * than an edge case.
 */

/** safeStorage is backed by a real OS secret store. */
const LEVEL_OS = 'os';
/** safeStorage "works" but the key is not secret (Linux basic_text). */
const LEVEL_WEAK = 'weak';
/** safeStorage cannot encrypt; secrets are written as plaintext. */
const LEVEL_NONE = 'none';

/**
 * @param {{available: boolean, backend?: string, platform?: string}} input
 *   `available` is isEncryptionAvailable(); `backend` is
 *   getSelectedStorageBackend() where supported, otherwise undefined.
 * @returns {{level: string, protected: boolean, backend: string|null, summary: string, detail: string}}
 */
function assessSecretStorage({ available, backend, platform } = {}) {
  const normalizedBackend = typeof backend === 'string' && backend ? backend : null;

  if (!available) {
    return {
      level: LEVEL_NONE,
      protected: false,
      backend: normalizedBackend,
      summary: 'Not encrypted',
      detail: platform === 'linux'
        ? 'No OS secret store is available, so the Cloudflare API token is stored as plain text. Install and unlock a keyring (gnome-keyring or kwallet), or clear the token and re-enter it each session.'
        : 'No OS secret store is available, so the Cloudflare API token is stored as plain text. Clear the token and re-enter it each session if this machine is shared.'
    };
  }

  if (normalizedBackend === 'basic_text') {
    return {
      level: LEVEL_WEAK,
      protected: false,
      backend: normalizedBackend,
      summary: 'Weakly encrypted',
      detail: 'No keyring was found, so the token is encrypted with a hardcoded password rather than a secret key. Treat it as readable by anyone with the file. Install and unlock gnome-keyring or kwallet to fix this.'
    };
  }

  return {
    level: LEVEL_OS,
    protected: true,
    backend: normalizedBackend,
    summary: 'Encrypted by the OS keystore',
    detail: normalizedBackend
      ? `Secrets are encrypted with a key held by ${normalizedBackend}.`
      : 'Secrets are encrypted with a key held by the OS keystore.'
  };
}

module.exports = {
  LEVEL_OS,
  LEVEL_WEAK,
  LEVEL_NONE,
  assessSecretStorage
};
