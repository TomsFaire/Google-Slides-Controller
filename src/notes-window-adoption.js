'use strict';

/**
 * Speaker-notes (presenter view) window adoption.
 *
 * Google Slides has no API to open presenter view, so the app fakes an "s"
 * keypress into the presentation window; Slides responds with window.open().
 * To control that popup — position it on the notes display, apply the notes
 * layout preference, read notes, zoom, scroll — the app must adopt it as
 * `notesWindow`.
 *
 * Identify the popup by its OPENER (the presentation webContents), not by its
 * URL. Google's presenter-view URL is undocumented and carries no stable
 * "speaker"/"notes" substring, so URL matching silently fails to adopt and
 * every downstream notes feature goes dead. Scoping the hook to the
 * presentation window's own `did-create-window` also keeps it from grabbing
 * unrelated app windows (key/fill outputs, overlays), which an app-level
 * `browser-window-created` listener could.
 *
 * The hook is one-shot: it disarms as soon as it adopts a window. Every path
 * that (re)triggers the "s" keypress must arm a fresh hook.
 *
 * @param {object} options
 * @param {import('electron').WebContents} options.webContents Presentation window webContents.
 * @param {(win: import('electron').BrowserWindow) => void} options.onAdopt Called once with the popup.
 * @returns {{ dispose: () => void, isArmed: () => boolean }}
 */
function armNotesAdoption({ webContents, onAdopt } = {}) {
  if (!webContents || typeof webContents.on !== 'function') {
    throw new TypeError('armNotesAdoption requires a webContents with .on()');
  }
  if (typeof onAdopt !== 'function') {
    throw new TypeError('armNotesAdoption requires an onAdopt callback');
  }

  let armed = true;

  function dispose() {
    if (!armed) return;
    armed = false;
    try {
      if (typeof webContents.isDestroyed !== 'function' || !webContents.isDestroyed()) {
        webContents.removeListener('did-create-window', listener);
      }
    } catch (e) {
      // webContents already torn down; nothing to detach.
    }
  }

  function listener(win) {
    if (!armed) return;
    dispose();
    onAdopt(win);
  }

  webContents.on('did-create-window', listener);

  return { dispose, isArmed: () => armed };
}

module.exports = { armNotesAdoption };
