/**
 * Electron support-window arithmetic.
 *
 * Electron supports the latest N majors (N = 3 as of this writing). Deliberately
 * pure: no I/O and no `require('electron')`, so it can be unit tested. Live data
 * fetching lives in scripts/check-electron-support.js.
 */

const DEFAULT_SUPPORTED_MAJORS = 3;

/**
 * Exit codes. Load-bearing: the drift workflow branches on these, so a
 * transient network blip must never be confusable with a real fault.
 *   0 OK               - within the support window, or drift that is not yet critical
 *   1 DRIFT_CRITICAL   - 2+ majors past end-of-life; fail the run
 *   2 NETWORK          - could not reach the registry; retry later, NOT drift
 *   3 INTERNAL         - bad config or a bug; needs a human, NOT retry-safe
 */
const EXIT = { OK: 0, DRIFT_CRITICAL: 1, NETWORK: 2, INTERNAL: 3 };

/** Tag an error as a network failure so the exit-code mapping can see it. */
function networkError(message) {
  const err = new Error(message);
  err.kind = 'network';
  return err;
}

/**
 * HTTP statuses worth retrying. Everything else in 4xx means the request itself is
 * wrong -- the endpoint moved, or we are not allowed -- which no amount of retrying
 * fixes. Treating those as transient is how an alarm goes quiet forever.
 */
const RETRYABLE_HTTP_STATUSES = new Set([408, 429]);

/**
 * Classify a non-OK HTTP response. Permanent 4xx failures are returned UNTAGGED so
 * they map to EXIT.INTERNAL and surface loudly; 5xx and rate-limit/timeout responses
 * are tagged as network failures and stay retry-safe.
 */
function classifyHttpFailure(url, status) {
  const message = `GET ${url} -> ${status}`;
  if (status >= 400 && status < 500 && !RETRYABLE_HTTP_STATUSES.has(status)) {
    return new Error(`${message} (permanent: endpoint moved or access denied)`);
  }
  return networkError(message);
}

/**
 * Map a thrown error to an exit code. Only errors explicitly tagged as network
 * failures get the retry-safe code; everything else is INTERNAL and must be loud.
 */
function exitCodeForFailure(err) {
  return err && err.kind === 'network' ? EXIT.NETWORK : EXIT.INTERNAL;
}

/** Extract the major version from a version or semver range. */
function parseMajor(versionRange) {
  const match = String(versionRange || '').match(/(\d+)\s*\./) || String(versionRange || '').match(/(\d+)\s*$/);
  if (!match) {
    throw new Error(`unparseable Electron version: ${JSON.stringify(versionRange)}`);
  }
  return Number(match[1]);
}

/**
 * Assess a major against the support window.
 *
 * States:
 *   supported   - within the latest N majors; callers should stay silent
 *   unsupported - exactly 1 major past end-of-life; open one issue
 *   critical    - 2 or more majors past end-of-life; escalate and fail the run
 */
function assessSupport({ currentMajor, latestMajor, supportedMajors = DEFAULT_SUPPORTED_MAJORS }) {
  for (const [name, value] of [['currentMajor', currentMajor], ['latestMajor', latestMajor]]) {
    if (typeof value !== 'number' || !Number.isFinite(value)) {
      throw new Error(`${name} must be a number, got ${JSON.stringify(value)}`);
    }
  }

  if (typeof supportedMajors !== 'number' || !Number.isInteger(supportedMajors) || supportedMajors < 1) {
    throw new Error(`supportedMajors must be a positive integer, got ${JSON.stringify(supportedMajors)}`);
  }

  const oldestSupportedMajor = latestMajor - (supportedMajors - 1);
  const majorsBehind = latestMajor - currentMajor;
  const majorsPastEol = Math.max(0, oldestSupportedMajor - currentMajor);

  let state = 'supported';
  if (majorsPastEol === 1) state = 'unsupported';
  else if (majorsPastEol >= 2) state = 'critical';

  return { state, majorsBehind, majorsPastEol, oldestSupportedMajor };
}

module.exports = {
  parseMajor,
  assessSupport,
  exitCodeForFailure,
  networkError,
  classifyHttpFailure,
  RETRYABLE_HTTP_STATUSES,
  EXIT,
  DEFAULT_SUPPORTED_MAJORS
};
