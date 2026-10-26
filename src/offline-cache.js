'use strict';

/**
 * Offline presentation cache — pure helpers extracted from main.js.
 *
 * The recording/replay machinery itself needs Electron (CDP debugger,
 * webRequest, BrowserWindow), but three pieces of it are plain logic that must
 * not be allowed to regress:
 *
 *   1. Reading a cache manifest off disk. A manifest written by an older
 *      build, truncated by a crash mid-write, or hand-edited can be valid JSON
 *      and still carry no usable `entries` map. Reading `.entries` straight off
 *      it throws or yields `undefined`, which leaves the active-manifest
 *      globals half-set and offline playback silently broken.
 *
 *   2. Serialising cache warm-ups. A warm attaches a CDP debugger, registers a
 *      'debugger-message' listener and drives the deck with key events; two of
 *      them at once orphan listeners and interleave recorded responses. A
 *      cache-state check is not a guard — state is set inside the async run, so
 *      overlapping callers can both get past it. The guard below keys off the
 *      in-flight promise itself.
 *
 *   3. Working out how far back to walk the deck once a warm finishes, so the
 *      presenter is returned to the slide they were on.
 */

/** Cache states broadcast to the desktop UI, the web remote and /api/status. */
const OFFLINE_CACHE_STATES = Object.freeze({
  NOT_CACHED: 'not-cached',
  CACHING: 'caching',
  CACHED: 'cached',
  OFFLINE: 'offline',
  ERROR: 'error',
});

/**
 * Validate a parsed cache manifest and return its usable contents.
 *
 * Returns `null` — never throws — for anything that cannot drive offline
 * playback: a missing/!object manifest, a missing or non-object `entries` map,
 * or an `entries` map with no entry carrying a `hash` to read a body file by.
 * Entries without a usable `hash` are dropped rather than failing the whole
 * manifest, so one bad resource does not discard a good cache.
 *
 * @param {unknown} manifest Parsed `cache-manifest.json` contents.
 * @returns {{ entries: Record<string, object>, cachedAt: string|null, slideCount: number|null }|null}
 */
function readManifestEntries(manifest) {
  if (!manifest || typeof manifest !== 'object' || Array.isArray(manifest)) return null;

  const rawEntries = manifest.entries;
  if (!rawEntries || typeof rawEntries !== 'object' || Array.isArray(rawEntries)) return null;

  const entries = {};
  for (const [url, entry] of Object.entries(rawEntries)) {
    if (!entry || typeof entry !== 'object') continue;
    if (typeof entry.hash !== 'string' || entry.hash === '') continue;
    entries[url] = entry;
  }
  if (Object.keys(entries).length === 0) return null;

  return {
    entries,
    cachedAt: typeof manifest.cachedAt === 'string' ? manifest.cachedAt : null,
    slideCount: Number.isFinite(manifest.slideCount) ? manifest.slideCount : null,
  };
}

/**
 * Reverse index for serving bodies over the `gs-cache://<hash>` protocol.
 *
 * @param {Record<string, { hash: string }>} entries url → entry
 * @returns {Record<string, string>} hash → url
 */
function buildHashToUrl(entries) {
  if (!entries || typeof entries !== 'object') return {};
  return Object.fromEntries(
    Object.entries(entries)
      .filter(([, entry]) => entry && typeof entry.hash === 'string' && entry.hash !== '')
      .map(([url, entry]) => [entry.hash, url])
  );
}

/**
 * Serialise cache warm-ups behind the in-flight promise.
 *
 * `start()` runs `runner` only when nothing is running, and reports which
 * happened so callers can answer an IPC/HTTP request honestly instead of
 * claiming a second warm started. The returned promise is the live run in both
 * cases, so a caller that wants to wait can `await` it. Rejections propagate —
 * callers attach their own handler and set the error cache state.
 *
 * @param {(...args: any[]) => Promise<any>} runner The warm implementation.
 * @returns {{ start: (...args: any[]) => { started: boolean, reason: string|null, promise: Promise<any> }, isWarming: () => boolean, whenIdle: () => Promise<any> }}
 */
function createWarmGuard(runner) {
  if (typeof runner !== 'function') {
    throw new TypeError('createWarmGuard requires a runner function');
  }

  let inFlight = null;

  function start(...args) {
    if (inFlight) {
      return { started: false, reason: 'already-warming', promise: inFlight };
    }

    // Deferred to a microtask so `inFlight` is assigned before `runner` can run
    // — a runner that throws synchronously must not leave the guard latched.
    const promise = Promise.resolve()
      .then(() => runner(...args))
      .finally(() => {
        if (inFlight === promise) inFlight = null;
      });

    inFlight = promise;
    return { started: true, reason: null, promise };
  }

  return {
    start,
    isWarming: () => inFlight !== null,
    whenIdle: () => (inFlight ? inFlight.catch(() => {}) : Promise.resolve()),
  };
}

/**
 * How many slides to walk back to return the presenter to where they started.
 *
 * Clamped at zero: a warm that ended level with — or behind — the starting
 * slide must not drive the deck anywhere. Non-numeric readings (the notes
 * window could not be read) mean "do not move".
 *
 * @param {number} startSlide 1-based slide the warm began on.
 * @param {number} endSlide 1-based slide the warm finished on.
 * @returns {number}
 */
function computeRestoreSteps(startSlide, endSlide) {
  if (!Number.isFinite(startSlide) || !Number.isFinite(endSlide)) return 0;
  return Math.max(0, Math.floor(endSlide) - Math.floor(startSlide));
}

module.exports = {
  OFFLINE_CACHE_STATES,
  readManifestEntries,
  buildHashToUrl,
  createWarmGuard,
  computeRestoreSteps,
};
