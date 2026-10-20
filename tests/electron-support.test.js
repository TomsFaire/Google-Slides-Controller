const { test } = require('node:test');
const assert = require('node:assert/strict');

const { parseMajor, assessSupport, exitCodeForFailure, networkError, classifyHttpFailure, EXIT } = require('../src/electron-support');

// ── parseMajor ───────────────────────────────────────────────────────────────

test('parseMajor reads the major from a caret range', () => {
  assert.equal(parseMajor('^44.4.5'), 44);
});

test('parseMajor reads the major from an exact version', () => {
  assert.equal(parseMajor('44.4.5'), 44);
});

test('parseMajor reads the major from tilde and gte ranges', () => {
  assert.equal(parseMajor('~44.0.0'), 44);
  assert.equal(parseMajor('>=44'), 44);
});

test('parseMajor throws on unparseable input', () => {
  assert.throws(() => parseMajor('not-a-version'), /unparseable/i);
  assert.throws(() => parseMajor(''), /unparseable/i);
});

// ── assessSupport: the healthy case must be silent ───────────────────────────

test('the latest major is supported', () => {
  const r = assessSupport({ currentMajor: 44, latestMajor: 44 });
  assert.equal(r.state, 'supported');
  assert.equal(r.majorsBehind, 0);
  assert.equal(r.majorsPastEol, 0);
});

test('the oldest of the latest three majors is still supported', () => {
  const r = assessSupport({ currentMajor: 42, latestMajor: 44 });
  assert.equal(r.state, 'supported');
  assert.equal(r.majorsBehind, 2);
  assert.equal(r.oldestSupportedMajor, 42);
});

// ── assessSupport: the escalation boundary ───────────────────────────────────

test('one major past end-of-life is unsupported, not yet critical', () => {
  const r = assessSupport({ currentMajor: 41, latestMajor: 44 });
  assert.equal(r.state, 'unsupported');
  assert.equal(r.majorsPastEol, 1);
  assert.equal(r.majorsBehind, 3);
});

test('two majors past end-of-life escalates to critical', () => {
  const r = assessSupport({ currentMajor: 40, latestMajor: 44 });
  assert.equal(r.state, 'critical');
  assert.equal(r.majorsPastEol, 2);
});

test('the situation this whole process exists to prevent', () => {
  // Electron 33 against a latest of 44 — where the app actually was.
  const r = assessSupport({ currentMajor: 33, latestMajor: 44 });
  assert.equal(r.state, 'critical');
  assert.equal(r.majorsBehind, 11);
  assert.equal(r.majorsPastEol, 9);
});

// ── assessSupport: edges ─────────────────────────────────────────────────────

test('running ahead of stable (alpha/beta) counts as supported', () => {
  const r = assessSupport({ currentMajor: 45, latestMajor: 44 });
  assert.equal(r.state, 'supported');
  assert.equal(r.majorsBehind, -1);
  assert.equal(r.majorsPastEol, 0);
});

test('the support window size is configurable', () => {
  const r = assessSupport({ currentMajor: 43, latestMajor: 44, supportedMajors: 1 });
  assert.equal(r.state, 'unsupported');
  assert.equal(r.oldestSupportedMajor, 44);
});

test('assessSupport rejects non-numeric input rather than guessing', () => {
  assert.throws(() => assessSupport({ currentMajor: NaN, latestMajor: 44 }), /must be a number/i);
  assert.throws(() => assessSupport({ currentMajor: 44, latestMajor: undefined }), /must be a number/i);
});

// ── exit-code contract: the workflow branches on these ───────────────────────

test('a tagged network failure maps to the retry-safe exit code', () => {
  assert.equal(exitCodeForFailure(networkError('registry unreachable')), EXIT.NETWORK);
  assert.equal(EXIT.NETWORK, 2);
});

test('an untagged error maps to INTERNAL, never to the retry-safe code', () => {
  assert.equal(exitCodeForFailure(new Error('unparseable Electron version: undefined')), EXIT.INTERNAL);
  assert.equal(exitCodeForFailure(undefined), EXIT.INTERNAL);
  assert.equal(EXIT.INTERNAL, 3);
  assert.notEqual(EXIT.INTERNAL, EXIT.NETWORK);
});

test('assessSupport rejects a nonsensical support-window size', () => {
  assert.throws(() => assessSupport({ currentMajor: 40, latestMajor: 44, supportedMajors: 0 }), /positive integer/i);
  assert.throws(() => assessSupport({ currentMajor: 40, latestMajor: 44, supportedMajors: 2.5 }), /positive integer/i);
});

// ── HTTP classification: a permanent failure must never look retry-safe ──────

test('a permanent 4xx is classified INTERNAL, so a moved endpoint cannot go quiet', () => {
  for (const status of [400, 401, 403, 404, 410, 451]) {
    const err = classifyHttpFailure('https://registry.example/x', status);
    assert.equal(exitCodeForFailure(err), EXIT.INTERNAL, `status ${status} must be INTERNAL`);
    assert.match(err.message, /permanent/);
  }
});

test('5xx and rate-limit/timeout responses stay retry-safe', () => {
  for (const status of [408, 429, 500, 502, 503, 504]) {
    assert.equal(exitCodeForFailure(classifyHttpFailure('https://registry.example/x', status)),
      EXIT.NETWORK, `status ${status} must be NETWORK`);
  }
});
