const { test } = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');

const { armNotesAdoption } = require('../src/notes-window-adoption');

// Minimal stand-in for an Electron webContents: an EventEmitter plus isDestroyed().
function fakeWebContents() {
  const wc = new EventEmitter();
  wc.isDestroyed = () => false;
  return wc;
}

// The real presenter-view popup URL. Note it contains neither "speakernotes"
// nor "speaker" — matching on those substrings is what broke adoption.
const PRESENTER_VIEW_URL =
  'https://docs.google.com/presentation/d/1qKhywpFhjG4tAtA1e2Rk9dB2lVk_uu5_Ol5TaBhvvPo/presentnotes?authuser=0&rm=minimal';

function fakeWindow(url) {
  return { webContents: { getURL: () => url } };
}

// ── adoption is opener-based, not URL-based ──────────────────────────────────

test('adopts a presenter-view popup whose URL contains no "speaker" substring', () => {
  const wc = fakeWebContents();
  const adopted = [];
  armNotesAdoption({ webContents: wc, onAdopt: win => adopted.push(win) });

  const popup = fakeWindow(PRESENTER_VIEW_URL);
  wc.emit('did-create-window', popup);

  assert.deepEqual(adopted, [popup]);
});

test('adopts a popup that has not navigated yet (empty URL at creation)', () => {
  const wc = fakeWebContents();
  const adopted = [];
  armNotesAdoption({ webContents: wc, onAdopt: win => adopted.push(win) });

  const popup = fakeWindow('');
  wc.emit('did-create-window', popup);

  assert.equal(adopted.length, 1);
});

// ── one-shot semantics ───────────────────────────────────────────────────────

test('adopts only the first popup, then disarms', () => {
  const wc = fakeWebContents();
  const adopted = [];
  const hook = armNotesAdoption({ webContents: wc, onAdopt: win => adopted.push(win) });

  const first = fakeWindow(PRESENTER_VIEW_URL);
  const second = fakeWindow(PRESENTER_VIEW_URL);
  wc.emit('did-create-window', first);
  wc.emit('did-create-window', second);

  assert.deepEqual(adopted, [first]);
  assert.equal(hook.isArmed(), false);
});

test('removes its listener from the webContents once adopted', () => {
  const wc = fakeWebContents();
  armNotesAdoption({ webContents: wc, onAdopt: () => {} });
  assert.equal(wc.listenerCount('did-create-window'), 1);

  wc.emit('did-create-window', fakeWindow(PRESENTER_VIEW_URL));

  assert.equal(wc.listenerCount('did-create-window'), 0);
});

// ── relaunch: re-arming must track the new popup ─────────────────────────────

test('re-arming after adoption tracks the relaunched notes popup', () => {
  const wc = fakeWebContents();
  const adopted = [];
  const onAdopt = win => adopted.push(win);

  armNotesAdoption({ webContents: wc, onAdopt });
  const first = fakeWindow(PRESENTER_VIEW_URL);
  wc.emit('did-create-window', first);

  // Relaunch: the notes window was closed and "s" pressed again.
  armNotesAdoption({ webContents: wc, onAdopt });
  const relaunched = fakeWindow(PRESENTER_VIEW_URL);
  wc.emit('did-create-window', relaunched);

  assert.deepEqual(adopted, [first, relaunched]);
});

// ── disposal ─────────────────────────────────────────────────────────────────

test('dispose() prevents adoption and detaches the listener', () => {
  const wc = fakeWebContents();
  const adopted = [];
  const hook = armNotesAdoption({ webContents: wc, onAdopt: win => adopted.push(win) });

  hook.dispose();
  wc.emit('did-create-window', fakeWindow(PRESENTER_VIEW_URL));

  assert.deepEqual(adopted, []);
  assert.equal(wc.listenerCount('did-create-window'), 0);
  assert.equal(hook.isArmed(), false);
});

test('dispose() is idempotent', () => {
  const wc = fakeWebContents();
  const hook = armNotesAdoption({ webContents: wc, onAdopt: () => {} });
  hook.dispose();
  assert.doesNotThrow(() => hook.dispose());
});

test('dispose() tolerates an already-destroyed webContents', () => {
  const wc = fakeWebContents();
  const hook = armNotesAdoption({ webContents: wc, onAdopt: () => {} });
  wc.isDestroyed = () => true;
  assert.doesNotThrow(() => hook.dispose());
});

// ── guards ───────────────────────────────────────────────────────────────────

test('throws when given no webContents', () => {
  assert.throws(() => armNotesAdoption({ onAdopt: () => {} }), TypeError);
});

test('throws when given no onAdopt callback', () => {
  assert.throws(() => armNotesAdoption({ webContents: fakeWebContents() }), TypeError);
});
