const { test } = require('node:test');
const assert = require('node:assert/strict');

const {
  OFFLINE_CACHE_STATES,
  readManifestEntries,
  buildHashToUrl,
  createWarmGuard,
  computeRestoreSteps,
} = require('../src/offline-cache');

// A manifest as warmOfflineCache writes it.
function validManifest() {
  return {
    cachedAt: '2026-06-01T10:00:00.000Z',
    slideCount: 12,
    entries: {
      'https://docs.google.com/presentation/d/abc/slide1.png': {
        hash: 'a'.repeat(64),
        mimeType: 'image/png',
        statusCode: 200,
        responseHeaders: {},
      },
      'https://fonts.gstatic.com/s/roboto.woff2': {
        hash: 'b'.repeat(64),
        mimeType: 'font/woff2',
        statusCode: 200,
        responseHeaders: {},
      },
    },
  };
}

// ── readManifestEntries — the null-guard ─────────────────────────────────────

test('reads entries, cachedAt and slideCount from a well-formed manifest', () => {
  const result = readManifestEntries(validManifest());

  assert.equal(Object.keys(result.entries).length, 2);
  assert.equal(result.cachedAt, '2026-06-01T10:00:00.000Z');
  assert.equal(result.slideCount, 12);
});

test('returns null instead of throwing for a manifest with no entries key', () => {
  // A version-mismatched manifest: valid JSON, no `entries`. Reading
  // `.entries` straight off it leaves the active-manifest globals undefined.
  assert.equal(readManifestEntries({ cachedAt: '2026-06-01T10:00:00.000Z' }), null);
});

test('returns null for a manifest whose entries is not a usable map', () => {
  assert.equal(readManifestEntries({ entries: null }), null);
  assert.equal(readManifestEntries({ entries: 'corrupted' }), null);
  assert.equal(readManifestEntries({ entries: [] }), null);
  assert.equal(readManifestEntries({ entries: {} }), null);
});

test('returns null for manifests that are not objects at all', () => {
  assert.equal(readManifestEntries(null), null);
  assert.equal(readManifestEntries(undefined), null);
  assert.equal(readManifestEntries('{"entries":{}}'), null);
  assert.equal(readManifestEntries([]), null);
  assert.equal(readManifestEntries(7), null);
});

test('drops entries with no usable hash but keeps the rest of the cache', () => {
  const manifest = validManifest();
  manifest.entries['https://docs.google.com/broken'] = { mimeType: 'image/png' };
  manifest.entries['https://docs.google.com/empty-hash'] = { hash: '' };
  manifest.entries['https://docs.google.com/null-entry'] = null;

  const result = readManifestEntries(manifest);

  assert.equal(Object.keys(result.entries).length, 2);
  assert.ok(!('https://docs.google.com/broken' in result.entries));
  assert.ok(!('https://docs.google.com/empty-hash' in result.entries));
  assert.ok(!('https://docs.google.com/null-entry' in result.entries));
});

test('returns null when every entry is unusable', () => {
  assert.equal(readManifestEntries({ entries: { 'https://x/y': { mimeType: 'image/png' } } }), null);
});

test('normalises a missing or non-string cachedAt to null', () => {
  const manifest = validManifest();
  delete manifest.cachedAt;
  assert.equal(readManifestEntries(manifest).cachedAt, null);

  const numeric = validManifest();
  numeric.cachedAt = 1717236000000;
  assert.equal(readManifestEntries(numeric).cachedAt, null);
});

test('normalises a non-numeric slideCount to null', () => {
  const manifest = validManifest();
  manifest.slideCount = 'twelve';
  assert.equal(readManifestEntries(manifest).slideCount, null);
});

// ── buildHashToUrl ───────────────────────────────────────────────────────────

test('builds the hash → url reverse index used by gs-cache://', () => {
  const { entries } = readManifestEntries(validManifest());
  const index = buildHashToUrl(entries);

  assert.equal(index['a'.repeat(64)], 'https://docs.google.com/presentation/d/abc/slide1.png');
  assert.equal(index['b'.repeat(64)], 'https://fonts.gstatic.com/s/roboto.woff2');
});

test('reverse index skips entries with no hash rather than keying on undefined', () => {
  const index = buildHashToUrl({
    'https://x/good': { hash: 'abc' },
    'https://x/bad': { mimeType: 'image/png' },
  });

  assert.deepEqual(Object.keys(index), ['abc']);
});

test('reverse index of nothing is an empty object', () => {
  assert.deepEqual(buildHashToUrl(null), {});
  assert.deepEqual(buildHashToUrl(undefined), {});
});

// ── createWarmGuard — the serialisation guard ────────────────────────────────

// A runner whose completion the test controls. The guard defers the runner to
// a microtask, so settling waits for the call to land first.
function deferredRunner() {
  const calls = [];
  let settle = null;

  const runner = (...args) => {
    calls.push(args);
    return new Promise((resolve, reject) => { settle = { resolve, reject }; });
  };

  async function waitForCall() {
    for (let i = 0; i < 100 && !settle; i++) await null;
    if (!settle) throw new Error('runner was never called');
  }

  return {
    runner,
    calls,
    resolve: async v => { await waitForCall(); settle.resolve(v); settle = null; },
    reject: async e => { await waitForCall(); settle.reject(e); settle = null; },
  };
}

test('runs the warm when nothing is in flight', async () => {
  const d = deferredRunner();
  const guard = createWarmGuard(d.runner);

  const first = guard.start('window', 'deck-1');

  assert.equal(first.started, true);
  assert.equal(first.reason, null);
  assert.equal(guard.isWarming(), true);

  d.resolve('done');
  assert.equal(await first.promise, 'done');
  assert.deepEqual(d.calls, [['window', 'deck-1']]);
});

test('refuses a second warm while the first is still running', async () => {
  const d = deferredRunner();
  const guard = createWarmGuard(d.runner);

  const first = guard.start('window', 'deck-1');
  const second = guard.start('window', 'deck-1');
  const third = guard.start('window', 'deck-1');

  assert.equal(second.started, false);
  assert.equal(second.reason, 'already-warming');
  assert.equal(third.started, false);

  // The runner ran exactly once, no matter how many taps arrived.
  d.resolve('done');
  await first.promise;
  assert.equal(d.calls.length, 1);
});

test('a refused caller is handed the in-flight run so it can wait on it', async () => {
  const d = deferredRunner();
  const guard = createWarmGuard(d.runner);

  const first = guard.start();
  const second = guard.start();

  assert.equal(second.promise, first.promise);

  d.resolve('done');
  assert.equal(await second.promise, 'done');
});

test('rapid taps in the same tick cannot both start a warm', async () => {
  // The bug this guards: cache state is set inside the async run, so a
  // state-only check lets two synchronous callers both get through.
  const d = deferredRunner();
  const guard = createWarmGuard(d.runner);

  const results = [guard.start(), guard.start(), guard.start(), guard.start()];

  assert.deepEqual(results.map(r => r.started), [true, false, false, false]);

  d.resolve();
  await results[0].promise;
  assert.equal(d.calls.length, 1);
});

test('releases the guard once the warm finishes, so a later warm can run', async () => {
  const d = deferredRunner();
  const guard = createWarmGuard(d.runner);

  const first = guard.start();
  d.resolve();
  await first.promise;

  assert.equal(guard.isWarming(), false);

  const second = guard.start();
  assert.equal(second.started, true);
  d.resolve();
  await second.promise;
  assert.equal(d.calls.length, 2);
});

test('releases the guard when the warm rejects, and propagates the rejection', async () => {
  const d = deferredRunner();
  const guard = createWarmGuard(d.runner);

  const first = guard.start();
  d.reject(new Error('debugger attach failed'));

  await assert.rejects(first.promise, /debugger attach failed/);
  assert.equal(guard.isWarming(), false);

  // A failed warm must not latch the guard shut — the user can retry.
  const retry = guard.start();
  assert.equal(retry.started, true);
  d.resolve();
  await retry.promise;
});

test('releases the guard when the runner throws synchronously', async () => {
  const guard = createWarmGuard(() => { throw new Error('no presentation window'); });

  const first = guard.start();
  await assert.rejects(first.promise, /no presentation window/);
  assert.equal(guard.isWarming(), false);
});

test('whenIdle resolves after the in-flight warm settles, including on failure', async () => {
  const d = deferredRunner();
  const guard = createWarmGuard(d.runner);

  guard.start().promise.catch(() => {});
  const idle = guard.whenIdle();
  d.reject(new Error('boom'));

  await idle; // must not reject
  assert.equal(guard.isWarming(), false);
});

test('whenIdle on an idle guard resolves immediately', async () => {
  const guard = createWarmGuard(async () => {});
  await guard.whenIdle();
  assert.equal(guard.isWarming(), false);
});

test('rejects a guard built without a runner', () => {
  assert.throws(() => createWarmGuard(), TypeError);
  assert.throws(() => createWarmGuard('nope'), TypeError);
});

// ── computeRestoreSteps — presenter slide restoration ────────────────────────

test('walks back exactly as far as the warm advanced', () => {
  assert.equal(computeRestoreSteps(1, 12), 11);
  assert.equal(computeRestoreSteps(7, 12), 5);
});

test('does not move the deck when the warm ended where it started', () => {
  assert.equal(computeRestoreSteps(5, 5), 0);
});

test('never walks forward if the end slide reads behind the start', () => {
  assert.equal(computeRestoreSteps(9, 4), 0);
});

test('does not move the deck when a slide position could not be read', () => {
  assert.equal(computeRestoreSteps(null, 12), 0);
  assert.equal(computeRestoreSteps(1, undefined), 0);
  assert.equal(computeRestoreSteps(NaN, 12), 0);
  assert.equal(computeRestoreSteps(1, Infinity), 0);
});

// ── cache states ─────────────────────────────────────────────────────────────

test('exposes an error state so a failed warm has somewhere to land', () => {
  assert.equal(OFFLINE_CACHE_STATES.ERROR, 'error');
  assert.notEqual(OFFLINE_CACHE_STATES.ERROR, OFFLINE_CACHE_STATES.CACHING);
});

test('cache states are frozen', () => {
  assert.equal(Object.isFrozen(OFFLINE_CACHE_STATES), true);
});
