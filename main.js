/**
 * Google Slides Opener
 * 
 * Copyright (c) 2026 TomsFaire and contributors
 * Based on https://github.com/nerif-tafu/gslide-opener
 * Original work by nerif-tafu
 * 
 * Licensed under the MIT License
 */

const { app, BrowserWindow, ipcMain, screen, session, dialog, nativeImage, crashReporter, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const http = require('http');
const https = require('https');
const crypto = require('crypto');
const { execSync, spawn } = require('child_process');
let nodePty = null;
try {
  nodePty = require('node-pty');
} catch (e) {
  // node-pty optional
}
const os = require('os');
const net = require('net');
const dns = require('dns');
const util = require('util');
const QRCode = require('qrcode');

// ----------------------------
// Logging helpers (secure by default)
// ----------------------------
// Verbose mode can be enabled either via preferences (`verboseLogging: true`)
// or via environment variables (useful before preferences exist).
const VERBOSE_ENV_ENABLED =
  String(process.env.GSLIDE_OPENER_VERBOSE || '').toLowerCase() === '1' ||
  String(process.env.GS_OPENER_VERBOSE || '').toLowerCase() === '1' ||
  String(process.env.DEBUG || '').toLowerCase() === '1';

let verboseLoggingEnabled = VERBOSE_ENV_ENABLED;

// Redact common secret fields in ANY logs (even verbose).
const SECRET_KEY_RE = /(api[\-_]?key|token|secret|password|passphrase|authorization)/i;

function safeStringify(value, space = 0) {
  try {
    return JSON.stringify(
      value,
      (k, v) => {
        if (k === 'webUiTunnelPinScrypt' || k === 'webUiTunnelPinSalt') {
          return v ? '[REDACTED]' : v;
        }
        if (k && SECRET_KEY_RE.test(String(k))) {
          return v ? '[REDACTED]' : v;
        }
        return v;
      },
      space
    );
  } catch (e) {
    return '[Unserializable]';
  }
}

function setVerboseLoggingFromPrefs(prefs) {
  if (!prefs || typeof prefs !== 'object') return;
  if (prefs.verboseLogging === true) {
    verboseLoggingEnabled = true;
  } else if (prefs.verboseLogging === false) {
    // Allow env var to force verbose on even if pref is off
    verboseLoggingEnabled = VERBOSE_ENV_ENABLED;
  }
}

function logDebug(...args) {
  if (!verboseLoggingEnabled) return;
  console.log(...args);
}
function logInfo(...args) {
  console.log(...args);
}
function logWarn(...args) {
  console.warn(...args);
}
function logError(...args) {
  console.error(...args);
}

// Count U+FFFD (replacement characters) in raw notes text before normalization.
function countReplacementChars(text) {
  if (text == null || typeof text !== 'string') return 0;
  const m = text.match(/\uFFFD/g);
  return m ? m.length : 0;
}

// Track last-seen encoding issue state so /api/status can expose it cheaply.
let lastNotesEncodingIssue = false;

// Normalize speaker-notes text: fix line breaks and replace corruption (U+FFFD) with newlines.
function normalizeSpeakerNotes(text) {
  if (text == null || typeof text !== 'string') return '';
  let s = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\u2028/g, '\n')
    .replace(/\u2029/g, '\n')
    .replace(/\uFFFD+/g, '\n')
    .replace(/\u0000/g, '');
  return s;
}

// Debug: log index and hex codes around the first U+FFFD in a string (to locate where corruption is introduced).
function logFirstReplacementCharContext(str, label) {
  if (str == null || typeof str !== 'string') return;
  const idx = str.indexOf('\uFFFD');
  if (idx < 0) return;
  const start = Math.max(0, idx - 10);
  const end = Math.min(str.length, idx + 11);
  const slice = str.slice(start, end);
  const hexCodes = Array.from(slice).map((c, i) => {
    const pos = start + i;
    const hex = 'U+' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0');
    return pos === idx ? '[' + hex + ']' : hex;
  }).join(' ');
  logDebug('[SpeakerNotes] ' + (label || '') + ' First U+FFFD at index ' + idx + ', hex around: ' + hexCodes);
}

// ----------------------------
// Live debug log capture (for desktop UI + export)
// ----------------------------
const LOG_BUFFER_MAX = 4000;
let logBuffer = [];

function sanitizeLogText(text) {
  let s = String(text ?? '');
  // Redact common key/value patterns in plain text logs
  s = s.replace(/(\b(api[\-_]?key|token|secret|password|passphrase|authorization)\b\s*[:=]\s*)([^\s,'"\\]+)/gi, '$1[REDACTED]');
  // Redact JSON style "apiKey":"..."
  s = s.replace(/("api[\-_]?key"\s*:\s*")[^"]*(")/gi, '$1[REDACTED]$2');
  s = s.replace(/("token"\s*:\s*")[^"]*(")/gi, '$1[REDACTED]$2');
  s = s.replace(/("secret"\s*:\s*")[^"]*(")/gi, '$1[REDACTED]$2');
  s = s.replace(/("password"\s*:\s*")[^"]*(")/gi, '$1[REDACTED]$2');
  return s;
}

function appendToLogBuffer(level, args) {
  try {
    const ts = new Date().toISOString();
    const msg = sanitizeLogText(util.format(...args));
    const line = `${ts} [${String(level).toUpperCase()}] ${msg}`;

    logBuffer.push(line);
    if (logBuffer.length > LOG_BUFFER_MAX) {
      logBuffer = logBuffer.slice(logBuffer.length - LOG_BUFFER_MAX);
    }

    if (mainWindow && !mainWindow.isDestroyed()) {
      try {
        mainWindow.webContents.send('app-log-line', line);
      } catch (e) {
        // ignore
      }
    }
  } catch (e) {
    // ignore failures
  }
}

// Patch console.* so anything logged by main process shows up in the UI/log export.
// Keep originals so we don't break Electron/Node expectations.
const _origConsoleLog = console.log.bind(console);
const _origConsoleWarn = console.warn.bind(console);
const _origConsoleError = console.error.bind(console);

console.log = (...args) => {
  _origConsoleLog(...args);
  appendToLogBuffer('log', args);
};
console.warn = (...args) => {
  _origConsoleWarn(...args);
  appendToLogBuffer('warn', args);
};
console.error = (...args) => {
  _origConsoleError(...args);
  appendToLogBuffer('error', args);
};

// ----------------------------
// Web UI favicon (use app icon)
// ----------------------------
let cachedFaviconPng = null;
let cachedFaviconDataUrl = null;

function getFaviconPngBuffer() {
  if (cachedFaviconPng) return cachedFaviconPng;

  // Prefer a .png if present; fall back to .icns (mac) and render as PNG
  const candidates = [
    path.join(__dirname, 'build', 'icon.png'),
    path.join(__dirname, 'build', 'icon.ico'),
    path.join(__dirname, 'build', 'icon.icns'),
    path.join(app.getAppPath ? app.getAppPath() : __dirname, 'build', 'icon.png'),
    path.join(app.getAppPath ? app.getAppPath() : __dirname, 'build', 'icon.ico'),
    path.join(app.getAppPath ? app.getAppPath() : __dirname, 'build', 'icon.icns'),
  ];

  for (const p of candidates) {
    try {
      if (!fs.existsSync(p)) continue;
      const img = nativeImage.createFromPath(p);
      if (!img || img.isEmpty()) continue;
      const resized = img.resize({ width: 32, height: 32, quality: 'good' });
      const png = resized.toPNG();
      if (png && png.length) {
        cachedFaviconPng = png;
        return cachedFaviconPng;
      }
    } catch (e) {
      // keep trying
    }
  }

  return null;
}

function getFaviconDataUrl() {
  if (cachedFaviconDataUrl) return cachedFaviconDataUrl;
  const png = getFaviconPngBuffer();
  if (!png) return null;
  cachedFaviconDataUrl = `data:image/png;base64,${png.toString('base64')}`;
  return cachedFaviconDataUrl;
}

// Cached build info (version/buildNumber) for status + UI strings
let appBuildInfo = { version: 'unknown', buildNumber: 'unknown' };
try {
  const packageJsonPath = path.join(__dirname, 'package.json');
  const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
  appBuildInfo = {
    version: packageJson.version || 'unknown',
    buildNumber: packageJson.buildNumber || 'unknown'
  };
} catch (error) {
  logError('[Build Info] Error loading package.json:', error.message);
}

// Start crash reporter (local-only; dumps written to crashDumps directory)
try {
  crashReporter.start({
    uploadToServer: false,
    extra: {
      version: appBuildInfo.version,
      buildNumber: String(appBuildInfo.buildNumber),
      platform: process.platform
    }
  });
} catch (e) {
  logError('[Crash] crashReporter.start failed:', e.message);
}

// Main process crash handlers: log, write report, then quit
function handleMainProcessCrash(label, err) {
  const message = err && (err.message || String(err));
  const stack = err && err.stack ? err.stack : '';
  const tail = logBuffer.slice(-LOG_TAIL_LINES).join('\n');
  const content = [
    `Google Slides Opener - ${label}`,
    `Time: ${new Date().toISOString()}`,
    `Version: ${appBuildInfo.version} Build: ${appBuildInfo.buildNumber}`,
    `Platform: ${process.platform}`,
    '',
    '--- Error ---',
    message,
    stack,
    '',
    '--- Last log lines ---',
    tail
  ].join('\n');
  const filePath = writeCrashReport('main', content);
  logError('[Crash]', message);
  if (filePath) {
    dialog.showMessageBox(null, {
      type: 'error',
      title: 'Application Error',
      message: 'The app encountered an error and will close.',
      detail: `A crash report was saved to:\n${filePath}`,
      buttons: ['OK']
    }).then(() => app.quit()).catch(() => app.quit());
  } else {
    app.quit();
  }
}

process.on('uncaughtException', (err) => {
  handleMainProcessCrash('Uncaught Exception', err);
});

process.on('unhandledRejection', (reason, promise) => {
  handleMainProcessCrash('Unhandled Rejection', reason instanceof Error ? reason : new Error(String(reason)));
});

let mainWindow;
let presentationWindow = null;
let notesWindow = null;
let notesNormalizeIntervalId = null;
let currentSlide = null; // best-effort: we track on our next/prev; DOM can override when notes window has aria-posinset/aria-setsize
let lastPresentationUrl = null; // Store the last-opened presentation URL for reload functionality

// Google Slides speaker notes zoom: discrete steps from native baseline (Zoom in / Zoom out toolbar; see /api/zoom-in-notes).
const NOTES_ZOOM_MIN_STEPS = -10;
const NOTES_ZOOM_MAX_STEPS = 40;
/** Tracked offset from Slides baseline; updated on API zoom and reload restore. */
let notesZoomStepsFromDefault = 0;

function clampNotesZoomSteps(n) {
  const x = Math.round(Number(n));
  if (Number.isNaN(x)) return 0;
  return Math.min(NOTES_ZOOM_MAX_STEPS, Math.max(NOTES_ZOOM_MIN_STEPS, x));
}

function getDefaultNotesZoomStepsFromPrefs() {
  const prefs = loadPreferences();
  const v = prefs.defaultNotesZoomSteps;
  if (v === undefined || v === null) return 0;
  return clampNotesZoomSteps(v);
}

function resetNotesZoomForNewPresentation() {
  notesZoomStepsFromDefault = getDefaultNotesZoomStepsFromPrefs();
}

const NOTES_ZOOM_CLICK_IN_JS = `
(function() {
  const zoomInButton = document.querySelector('[title="Zoom in"]');
  if (!zoomInButton) return { success: false, error: 'Button not found' };
  const mousedownEvent = new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window, button: 0 });
  const mouseupEvent = new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window, button: 0 });
  const clickEvent = new MouseEvent('click', { bubbles: true, cancelable: true, view: window, button: 0 });
  zoomInButton.dispatchEvent(mousedownEvent);
  zoomInButton.dispatchEvent(mouseupEvent);
  zoomInButton.dispatchEvent(clickEvent);
  return { success: true };
})()
`.trim();

const NOTES_ZOOM_CLICK_OUT_JS = `
(function() {
  const zoomOutButton = document.querySelector('[title="Zoom out"]');
  if (!zoomOutButton) return { success: false, error: 'Button not found' };
  const mousedownEvent = new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window, button: 0 });
  const mouseupEvent = new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window, button: 0 });
  const clickEvent = new MouseEvent('click', { bubbles: true, cancelable: true, view: window, button: 0 });
  zoomOutButton.dispatchEvent(mousedownEvent);
  zoomOutButton.dispatchEvent(mouseupEvent);
  zoomOutButton.dispatchEvent(clickEvent);
  return { success: true };
})()
`.trim();

async function executeNotesZoomClickIn(notesWin) {
  if (!notesWin || notesWin.isDestroyed()) return { success: false, error: 'No notes window' };
  return notesWin.webContents.executeJavaScript(NOTES_ZOOM_CLICK_IN_JS);
}

async function executeNotesZoomClickOut(notesWin) {
  if (!notesWin || notesWin.isDestroyed()) return { success: false, error: 'No notes window' };
  return notesWin.webContents.executeJavaScript(NOTES_ZOOM_CLICK_OUT_JS);
}

async function applyNotesZoomSteps(notesWin, steps) {
  const n = clampNotesZoomSteps(steps);
  if (n === 0 || !notesWin || notesWin.isDestroyed()) return;
  const count = Math.abs(n);
  const inward = n > 0;
  for (let i = 0; i < count; i++) {
    if (notesWin.isDestroyed()) return;
    const r = inward ? await executeNotesZoomClickIn(notesWin) : await executeNotesZoomClickOut(notesWin);
    if (!r || !r.success) {
      logWarn('[Notes] Zoom replay stopped early at step', i + 1, '/', count);
      break;
    }
    if (i < count - 1) {
      await new Promise(resolve => setTimeout(resolve, 100));
    }
  }
}

async function applyTrackedNotesZoomWhenReady(win) {
  if (!win || win.isDestroyed()) return;
  const steps = notesZoomStepsFromDefault;
  if (steps === 0) return;
  for (let t = 0; t < 40; t++) {
    if (win.isDestroyed()) return;
    const ready = await win.webContents
      .executeJavaScript(`(function(){ return !!document.querySelector('[title="Zoom in"]'); })()`)
      .catch(() => false);
    if (ready) {
      await applyNotesZoomSteps(win, steps);
      logInfo('[Notes] Applied speaker notes zoom steps:', steps);
      return;
    }
    await new Promise(resolve => setTimeout(resolve, 150));
  }
  logWarn('[Notes] Zoom controls not ready; could not apply steps:', steps);
}

/** Optional: log native notes body font size before reload (DOM-only zoom cannot be mapped to steps reliably). */
async function logNotesZoomFontProbeBeforeReload(notesWin, label) {
  if (!notesWin || notesWin.isDestroyed()) return;
  try {
    const px = await notesWin.webContents.executeJavaScript(`
      (function() {
        var el = document.querySelector('div.punch-viewer-speakernotes-text-body-scrollable');
        if (!el) return null;
        var v = parseFloat(window.getComputedStyle(el).fontSize);
        return isNaN(v) ? null : v;
      })()
    `);
    if (px != null) logDebug('[Notes] Pre-reload zoom font probe' + (label ? ' ' + label : '') + ':', px, 'px (informational)');
  } catch (e) {
    logDebug('[Notes] Font probe skipped:', e.message);
  }
}

// ----------------------------
// Crash reporting and recovery
// ----------------------------
const CRASH_REPORTS_SUBDIR = 'crash-reports';
const LOG_TAIL_LINES = 500;

function getCrashReportsDir() {
  return path.join(app.getPath('userData'), CRASH_REPORTS_SUBDIR);
}

let lastCrashPath = null;
let lastCrashTime = null;

function writeCrashReport(type, content) {
  try {
    const dir = getCrashReportsDir();
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    const ts = new Date().toISOString().replace(/[:.]/g, '-');
    const name = `crash-${type}-${ts}.txt`;
    const filePath = path.join(dir, name);
    fs.writeFileSync(filePath, content, 'utf8');
    lastCrashPath = filePath;
    lastCrashTime = new Date().toISOString();
    return filePath;
  } catch (e) {
    logError('[Crash] Failed to write crash report:', e.message);
    return null;
  }
}

// Optional slideNumber: if provided and > 0, append #slide=id.pN so presentation opens on that slide
function toPresentUrl(inputUrl, slideNumber) {
  try {
    const u = new URL(inputUrl);
    const m = u.pathname.match(/\/presentation\/d\/([^/]+)/);
    if (!m) return inputUrl;
    const id = m[1];
    let base = `https://docs.google.com/presentation/d/${id}/present`;
    if (typeof slideNumber === 'number' && slideNumber > 0) {
      base += `#slide=id.p${slideNumber}`;
    }
    return base;
  } catch (e) {
    return inputUrl;
  }
}

// Use a persistent session for Google authentication
const GOOGLE_SESSION_PARTITION = 'persist:google';

// Build BrowserWindow options for the speaker notes popup so it opens at notes-display size.
// Google Slides presenter view uses a responsive layout: wide viewport = narrow preview column + wide notes; small viewport = 50/50. Opening at full size avoids the 50/50 split.
function getSpeakerNotesWindowOptions(notesDisplay) {
  const bounds = notesDisplay && notesDisplay.bounds ? notesDisplay.bounds : screen.getPrimaryDisplay().bounds;
  return {
    frame: false,
    x: bounds.x,
    y: bounds.y,
    width: bounds.width,
    height: bounds.height,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      partition: GOOGLE_SESSION_PARTITION
    }
  };
}

const VALID_PRESENTATION_GPU_MODES = new Set(['default', 'disable-gpu', 'angle-metal', 'angle-gl', 'swiftshader']);

/** BrowserWindow constructor options shared by all Google Slides presentation windows. */
function getPresentationBrowserWindowOptions(presentationBounds) {
  const b = presentationBounds && presentationBounds.width ? presentationBounds : screen.getPrimaryDisplay().bounds;
  return {
    x: b.x,
    y: b.y,
    width: b.width,
    height: b.height,
    frame: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      partition: GOOGLE_SESSION_PARTITION
    }
  };
}

function readPreferencesFileJsonSync() {
  try {
    const prefsPath = path.join(app.getPath('userData'), 'preferences.json');
    if (!fs.existsSync(prefsPath)) return {};
    return JSON.parse(fs.readFileSync(prefsPath, 'utf8'));
  } catch (e) {
    return {};
  }
}

/**
 * Apply GPU / ANGLE switches before app.ready(). Must stay in sync with loadPreferences() normalization.
 * Changing presentationGpuMode requires an app restart to take effect.
 */
function applyPresentationGpuCommandLineEarly() {
  let mode = 'default';
  try {
    const prefs = readPreferencesFileJsonSync();
    const raw = prefs && prefs.presentationGpuMode;
    if (raw && VALID_PRESENTATION_GPU_MODES.has(String(raw))) {
      mode = String(raw);
    }
  } catch (e) {
    mode = 'default';
  }
  if (mode === 'disable-gpu') {
    app.disableHardwareAcceleration();
  } else if (mode === 'angle-metal') {
    app.commandLine.appendSwitch('use-angle', 'metal');
  } else if (mode === 'angle-gl') {
    app.commandLine.appendSwitch('use-angle', 'gl');
  } else if (mode === 'swiftshader') {
    app.commandLine.appendSwitch('use-gl', 'swiftshader');
  }
}

applyPresentationGpuCommandLineEarly();

function wantsPresentationNativeFullscreen(prefs) {
  return process.platform === 'darwin' && prefs && prefs.presentationNativeFullscreen === true;
}

/** macOS: simple fullscreen (default) vs native fullscreen (diagnostic / compositor path). No-op on other OSes. */
function applyPresentationFullscreenChrome(presentationWindow, prefs) {
  if (process.platform !== 'darwin' || !presentationWindow || presentationWindow.isDestroyed()) return;
  const p = prefs || loadPreferences();
  try {
    if (wantsPresentationNativeFullscreen(p)) {
      presentationWindow.setFullScreen(true);
    } else {
      presentationWindow.setSimpleFullScreen(true);
    }
  } catch (e) {
    try {
      presentationWindow.setSimpleFullScreen(true);
    } catch (e2) {
      // ignore
    }
  }
}

function presentationFullscreenNeedsReapply(presentationWindow, prefs) {
  if (process.platform !== 'darwin' || !presentationWindow || presentationWindow.isDestroyed()) return false;
  const p = prefs || loadPreferences();
  if (wantsPresentationNativeFullscreen(p)) {
    return !presentationWindow.isFullScreen();
  }
  return !presentationWindow.isSimpleFullScreen();
}

// Normalize speaker notes text in the notes window: replace U+FFFD and "-" with real newlines in text nodes only.
// Uses TreeWalker so we don't replace the whole div (which broke slide updates). Runs on an interval.
// Set white-space: pre-wrap on the notes container so newline characters in text nodes actually render as line breaks.
function getNotesWindowNormalizeScript() {
  return `
(function(){
  var el = document.querySelector('div.punch-viewer-speakernotes-text-body-scrollable');
  if (!el) return;
  el.style.whiteSpace = 'pre-wrap';
  var nl = String.fromCharCode(10);
  var hyphens = { 45:1, 8208:1, 8209:1, 8210:1, 8211:1, 8212:1, 8213:1, 8738:1 };
  function fixText(str) {
    if (!str) return str;
    var out = '';
    for (var i = 0; i < str.length; i++) {
      var code = str.charCodeAt(i);
      if (code === 65533) { out += nl; if (i + 1 < str.length && hyphens[str.charCodeAt(i+1)]) i++; }
      else if (code === 65532) out += nl;
      else out += str[i];
    }
    return out;
  }
  var walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null, false);
  var node;
  while ((node = walker.nextNode())) {
    var fixed = fixText(node.nodeValue);
    if (fixed !== node.nodeValue) node.nodeValue = fixed;
  }
})();
  `.trim();
}

function normalizeNotesWindowContent() {
  if (!notesWindow || notesWindow.isDestroyed()) return;
  notesWindow.webContents.executeJavaScript(getNotesWindowNormalizeScript()).catch(() => {});
}

function startNotesWindowNormalizationInterval() {
  if (notesNormalizeIntervalId != null) return;
  notesNormalizeIntervalId = setInterval(normalizeNotesWindowContent, 1500);
}

function stopNotesWindowNormalizationInterval() {
  if (notesNormalizeIntervalId != null) {
    clearInterval(notesNormalizeIntervalId);
    notesNormalizeIntervalId = null;
  }
}

// CSS to hide the slide preview side panel in Google Slides presenter view,
// giving full width to the speaker notes text.
const PRESENTER_VIEW_HIDE_SIDE_PANEL_CSS = `
  /* Hide the slide preview panel and dragger */
  td.punch-viewer-speakernotes-side-panel {
    display: none !important;
  }
  div.punch-viewer-speakernotes-dragger {
    display: none !important;
  }

  /* Notes text takes full width */
  div.punch-viewer-speakernotes-text-body-scrollable {
    left: 0 !important;
  }
  div.punch-viewer-speakernotes-text-header-container {
    left: 0 !important;
  }
  div.punch-viewer-speaker-qanda-content {
    left: 0 !important;
  }
`;

function onNotesWindowCreated(win) {
  startNotesWindowNormalizationInterval();
  win.once('closed', stopNotesWindowNormalizationInterval);

  // Fix presenter view layout based on preference:
  // - 'hide': hide side panel entirely, notes get full width
  // - 'default': leave Google's layout as-is (50/50)
  win.webContents.on('did-finish-load', () => {
    if (win.isDestroyed()) return;
    const prefs = loadPreferences();
    const mode = prefs.notesLayout || 'hide';

    if (mode === 'hide') {
      win.webContents.insertCSS(PRESENTER_VIEW_HIDE_SIDE_PANEL_CSS).then(() => {
        logInfo('[Notes] Presenter layout: side panel hidden');
      }).catch(() => {});
    } else {
      logInfo('[Notes] Presenter layout: using Google default');
    }

    if (!win._gslideNotesZoomApplied) {
      win._gslideNotesZoomApplied = true;
      setTimeout(() => {
        applyTrackedNotesZoomWhenReady(win).catch(() => {});
      }, 450);
    }
  });

  // Persist notes window bounds to preferences on resize/move so they survive app restarts.
  let notesBoundsSaveTimer = null;
  const saveNotesBounds = () => {
    if (win.isDestroyed()) return;
    clearTimeout(notesBoundsSaveTimer);
    notesBoundsSaveTimer = setTimeout(() => {
      if (win.isDestroyed()) return;
      try {
        const bounds = win.getBounds();
        if (bounds.width > 100 && bounds.height > 100) {
          const prefs = loadPreferences();
          prefs.notesBounds = bounds;
          savePreferences(prefs);
          logDebug('[Notes] Saved notes window bounds to preferences:', bounds);
        }
      } catch (e) {
        logError('[Notes] Error saving notes bounds:', e);
      }
    }, 500);
  };
  win.on('resize', saveNotesBounds);
  win.on('move', saveNotesBounds);
  win.once('closed', () => clearTimeout(notesBoundsSaveTimer));
}

// Load saved notes window bounds from preferences (returns bounds object or null).
function loadSavedNotesBounds() {
  try {
    const prefs = loadPreferences();
    const b = prefs.notesBounds;
    if (b && b.width > 100 && b.height > 100) return b;
  } catch (e) {}
  return null;
}

/** Close presenter notes (if open) and reopen via the same shortcut as the HTTP API. */
async function relaunchSpeakerNotesWindow() {
  if (!presentationWindow || presentationWindow.isDestroyed()) {
    return { success: false, error: 'No presentation is open' };
  }
  console.log('[API] Relaunching speaker notes');
  if (notesWindow && !notesWindow.isDestroyed()) {
    notesWindow.close();
    notesWindow = null;
    await new Promise(resolve => setTimeout(resolve, 300));
  }
  presentationWindow.focus();
  await new Promise(resolve => setTimeout(resolve, 50));
  presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
  presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
  presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
  sendToBackups('/api/relaunch-speaker-notes', {}).catch(err => {
    console.error('[Backup] Error broadcasting relaunch-speaker-notes:', err);
  });
  return { success: true, message: 'Speaker notes relaunched' };
}

// Function to set speaker notes window to fullscreen on a specific display (preferred).
// Without targetDisplay, uses getDisplayMatching(window.getBounds()) which can follow the
// presentation window if the popup opened on the wrong monitor.
function setSpeakerNotesFullscreen(window, targetDisplay) {
  if (!window || window.isDestroyed()) return;
  // Make the window invisible (but NOT hidden) to prevent seeing the resize.
  // Using setOpacity(0) instead of hide() so the window keeps its full viewport dimensions.
  // Google Slides presenter view picks narrow-preview vs 50/50 layout based on viewport width
  // at render time; hide() collapses the viewport to 0x0, causing the wrong layout.
  window.setOpacity(0);

  // Wait for page to fully load before showing
  const showFullscreen = () => {
    if (window.isDestroyed()) return;

    try {
      let display;
      if (targetDisplay && targetDisplay.bounds) {
        display = targetDisplay;
      } else {
        try {
          const bounds = window.getBounds();
          if (bounds.width > 0 && bounds.height > 0) {
            display = screen.getDisplayMatching(bounds);
          } else {
            display = screen.getPrimaryDisplay();
          }
        } catch (e) {
          display = screen.getPrimaryDisplay();
        }
      }

      window.setBounds(display.bounds);
      if (process.platform === 'darwin') {
        // Use setSimpleFullScreen instead of setFullScreen to avoid creating a new Space
        // This prevents window management conflicts when "Displays have separate Spaces" is enabled
        window.setSimpleFullScreen(true);
      }
      window.setOpacity(1);
      window.show();
      logInfo('[Notes] Set speaker notes window to fullscreen (simple fullscreen to avoid Spaces conflicts)');
      // Fire resize events so Slides recalculates layout for the actual viewport width.
      const triggerResize = () => {
        if (window.isDestroyed()) return;
        window.webContents.executeJavaScript(`
          window.dispatchEvent(new Event('resize'));
          window.dispatchEvent(new UIEvent('resize'));
        `).catch(() => {});
      };
      setTimeout(triggerResize, 300);
      setTimeout(triggerResize, 800);
      setTimeout(triggerResize, 1500);
    } catch (error) {
      logError('[Notes] Error setting fullscreen:', error);
      // Fallback: just show the window
      if (!window.isDestroyed()) {
        window.setOpacity(1);
        window.show();
      }
    }
  };
  
  // Wait for page to finish loading, then wait a bit more for layout to stabilize
  window.webContents.once('did-finish-load', () => {
    setTimeout(showFullscreen, 1500);
  });
  if (window.webContents.isLoading() === false) {
    setTimeout(showFullscreen, 1500);
  } else {
    window.webContents.once('dom-ready', () => {
      setTimeout(showFullscreen, 1500);
    });
  }
}

// Apply cached bounds to speaker notes window (e.g. after reload so user's size is preserved)
function setSpeakerNotesBoundsFromCache(window, bounds) {
  if (!window || window.isDestroyed() || !bounds) return;
  window.setOpacity(0);
  const apply = () => {
    if (window.isDestroyed()) return;
    try {
      window.setBounds(bounds);
      window.setOpacity(1);
      window.show();
      logInfo('[Notes] Restored speaker notes window to cached size/position');
      // Trigger resize so Google Slides recalculates presenter view layout for actual viewport
      const triggerResize = () => {
        if (window.isDestroyed()) return;
        window.webContents.executeJavaScript(`
          window.dispatchEvent(new Event('resize'));
          window.dispatchEvent(new UIEvent('resize'));
        `).catch(() => {});
      };
      setTimeout(triggerResize, 300);
      setTimeout(triggerResize, 800);
      setTimeout(triggerResize, 1500);
    } catch (e) {
      logError('[Notes] Error restoring notes bounds:', e);
      if (!window.isDestroyed()) { window.setOpacity(1); window.show(); }
    }
  };
  window.webContents.once('did-finish-load', () => setTimeout(apply, 300));
  if (window.webContents.isLoading() === false) {
    setTimeout(apply, 300);
  } else {
    window.webContents.once('dom-ready', () => setTimeout(apply, 300));
  }
}

/** Display to use for speaker notes when prefs/API resolved a notes monitor (fallback: primary). */
function resolveNotesTargetDisplay(notesDisplay) {
  if (notesDisplay && notesDisplay.bounds) return notesDisplay;
  return screen.getPrimaryDisplay();
}

function boundsCenterIntersectsDisplay(bounds, display) {
  if (!bounds || !display || !display.bounds) return false;
  const cx = bounds.x + Math.floor(bounds.width / 2);
  const cy = bounds.y + Math.floor(bounds.height / 2);
  const b = display.bounds;
  return cx >= b.x && cx < b.x + b.width && cy >= b.y && cy < b.y + b.height;
}

/**
 * Restore saved notes size/position only if it lies on the selected notes display; otherwise
 * fullscreen the notes window on that display (avoids stale notesBounds pinning notes to the projector).
 */
function applySpeakerNotesInitialGeometry(window, notesDisplay, overrideBounds) {
  const target = resolveNotesTargetDisplay(notesDisplay);
  const b =
    overrideBounds && overrideBounds.width > 0 && overrideBounds.height > 0
      ? overrideBounds
      : loadSavedNotesBounds();
  if (b && b.width > 100 && b.height > 100 && boundsCenterIntersectsDisplay(b, target)) {
    setSpeakerNotesBoundsFromCache(window, b);
  } else {
    if (b && b.width > 100 && b.height > 100 && !boundsCenterIntersectsDisplay(b, target)) {
      logDebug(
        '[Notes] Cached notes bounds are not on the selected notes display; fullscreening on notes monitor instead'
      );
    }
    window.webContents.once('did-finish-load', () => setSpeakerNotesFullscreen(window, target));
    window.webContents.once('dom-ready', () => setTimeout(() => setSpeakerNotesFullscreen(window, target), 500));
  }
}

// Capture "current slide" and "next slide" preview images from the speaker notes (Presenter View) window.
// This avoids relying on <img src> URLs (which may be blob: and not accessible to remote devices).
async function captureSlidePreviewsFromNotesWindow({ maxSize = 200 } = {}) {
  if (!notesWindow || notesWindow.isDestroyed()) {
    return { success: false, error: 'No speaker notes window is open' };
  }

  // Find rectangles of the current/next slide preview elements inside the notes window.
  const rectInfo = await notesWindow.webContents.executeJavaScript(`
    (function () {
      function isVisible(el) {
        if (!el) return false;
        const style = window.getComputedStyle(el);
        if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return false;
        const r = el.getBoundingClientRect();
        if (!r || r.width < 60 || r.height < 60) return false;
        // Must intersect viewport
        if (r.bottom < 0 || r.right < 0 || r.top > window.innerHeight || r.left > window.innerWidth) return false;
        return true;
      }

      function rectOf(el) {
        const r = el.getBoundingClientRect();
        // Add a tiny padding to reduce border clipping
        const pad = 2;
        return {
          x: Math.max(0, Math.floor(r.left + pad)),
          y: Math.max(0, Math.floor(r.top + pad)),
          width: Math.max(1, Math.floor(r.width - (pad * 2))),
          height: Math.max(1, Math.floor(r.height - (pad * 2)))
        };
      }

      // Try known-ish presenter-view selectors first (best effort).
      const known = [];
      const knownSelectors = [
        // These may or may not exist depending on Slides updates
        '[aria-label*="Current slide"] img',
        '[aria-label*="Next slide"] img',
        '[aria-label*="Next slide"] canvas',
        '[aria-label*="Next"] img',
        'div[class*="punch-viewer"] img',
        'canvas',
        'iframe',
        'svg'
      ];

      for (const sel of knownSelectors) {
        try {
          document.querySelectorAll(sel).forEach(el => {
            if (isVisible(el)) known.push(el);
          });
        } catch (e) {}
      }

      // Broader fallback: look for large visible elements that likely render slide previews.
      const candidates = [];

      // Visible images/canvases/iframes/svgs (Slides often uses iframes/canvas/svg)
      document.querySelectorAll('img, canvas, iframe, svg').forEach(el => {
        if (!isVisible(el)) return;
        const r = el.getBoundingClientRect();
        const ar = r.width / Math.max(1, r.height);
        // Prefer slide-like aspect ratios (4:3 to 16:9-ish), but don't hard-reject yet
        candidates.push({ el, area: r.width * r.height, top: r.top, left: r.left });
      });

      // Visible divs with background-image (common for thumbnails)
      document.querySelectorAll('div[class*="punch-viewer"], div[style]').forEach(el => {
        if (!isVisible(el)) return;
        const style = window.getComputedStyle(el);
        const bg = style.backgroundImage || '';
        if (!bg || bg === 'none') return;
        const r = el.getBoundingClientRect();
        candidates.push({ el, area: r.width * r.height, top: r.top, left: r.left });
      });

      // Prefer left-side content (presenter preview pane is typically on the left)
      const midX = window.innerWidth / 2;
      const scored = candidates
        .map(c => {
          const r = c.el.getBoundingClientRect();
          const centerX = r.left + r.width / 2;
          const leftBias = centerX < midX ? 1.35 : 0.85;
          const topBias = r.top < 80 ? 0.5 : 1.0; // avoid grabbing header UI
          const ar = r.width / Math.max(1, r.height);
          // Slides are commonly 4:3 (1.33) or 16:9 (1.78)
          const arPenalty = Math.min(Math.abs(ar - 1.33), Math.abs(ar - 1.78));
          const arBias = (ar > 1.1 && ar < 2.1) ? (1.15 - Math.min(0.6, arPenalty)) : 0.65;
          return { ...c, score: c.area * leftBias * topBias * arBias, rect: r };
        })
        .sort((a, b) => b.score - a.score);

      // Pick the top 2 distinct elements (by rect separation)
      const picked = [];
      for (const item of scored) {
        if (picked.length >= 2) break;
        const r = item.rect;
        const overlapsTooMuch = picked.some(p => {
          const pr = p.rect;
          const overlapX = Math.max(0, Math.min(r.right, pr.right) - Math.max(r.left, pr.left));
          const overlapY = Math.max(0, Math.min(r.bottom, pr.bottom) - Math.max(r.top, pr.top));
          const overlapArea = overlapX * overlapY;
          return overlapArea > (Math.min(r.width * r.height, pr.width * pr.height) * 0.5);
        });
        if (!overlapsTooMuch) picked.push(item);
      }

      // If we didn't find enough, try using known list
      if (picked.length < 2 && known.length >= 2) {
        const k = Array.from(new Set(known)).filter(isVisible).map(el => {
          const r = el.getBoundingClientRect();
          return { el, rect: r, score: (r.width * r.height) * 1.1 };
        }).sort((a, b) => b.score - a.score);
        while (picked.length < 2 && k.length) picked.push(k.shift());
      }

      // Fallback: try to anchor the "Next" thumbnail by label text, then pick a large "current" preview above it.
      if (picked.length < 2) {
        function findByLabelText(txt) {
          const els = Array.from(document.querySelectorAll('*'))
            .filter(el => {
              if (!isVisible(el)) return false;
              const t = (el.textContent || '').trim();
              return t === txt;
            });
          // Prefer ones on the left half
          els.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
          return els[0] || null;
        }

        const nextLabel = findByLabelText('Next');
        if (nextLabel) {
          let container = nextLabel;
          for (let i = 0; i < 6 && container; i++) {
            const hasPreview = container.querySelector && container.querySelector('img,canvas,iframe,svg,div[style]');
            if (hasPreview) break;
            container = container.parentElement;
          }
          const nextPreviewEl = container ? (container.querySelector('img,canvas,iframe,svg') || container) : null;
          if (nextPreviewEl && isVisible(nextPreviewEl)) {
            const nextRect = nextPreviewEl.getBoundingClientRect();
            // Find a "current" preview above it: biggest slide-like element above nextRect.top
            const above = scored
              .filter(s => s.rect && (s.rect.top + s.rect.height) < (nextRect.top + 20))
              .sort((a, b) => b.score - a.score);
            if (above.length) {
              picked.push(above[0]);
              picked.push({ el: nextPreviewEl, rect: nextRect, score: nextRect.width * nextRect.height });
            }
          }
        }
      }

      if (picked.length < 2) {
        return { ok: false, error: 'Could not locate slide preview elements in presenter view' };
      }

      // Sort by vertical position: top = current, bottom = next
      picked.sort((a, b) => a.rect.top - b.rect.top);

      // Slide numbers from aria-posinset/aria-setsize when available
      let currentSlide = null;
      let totalSlides = null;
      try {
        const el = document.querySelector('[aria-posinset]');
        if (el) {
          const cur = parseInt(el.getAttribute('aria-posinset'), 10);
          const tot = parseInt(el.getAttribute('aria-setsize'), 10);
          if (!isNaN(cur)) currentSlide = cur;
          if (!isNaN(tot)) totalSlides = tot;
        }
      } catch (e) {}

      return {
        ok: true,
        current: rectOf(picked[0].el),
        next: rectOf(picked[1].el),
        currentSlide,
        totalSlides
      };
    })()
  `);

  if (!rectInfo || !rectInfo.ok || !rectInfo.current || !rectInfo.next) {
    return { success: false, error: rectInfo?.error || 'Failed to locate preview rectangles' };
  }

  function resizeToFit(nativeImg) {
    const size = nativeImg.getSize();
    const w = size.width || 1;
    const h = size.height || 1;
    const scale = Math.min(1, maxSize / Math.max(w, h));
    const targetW = Math.max(1, Math.round(w * scale));
    const targetH = Math.max(1, Math.round(h * scale));
    return nativeImg.resize({ width: targetW, height: targetH, quality: 'good' });
  }

  const currentImg = resizeToFit(await notesWindow.webContents.capturePage(rectInfo.current));
  const nextImg = resizeToFit(await notesWindow.webContents.capturePage(rectInfo.next));

  const currentDataUrl = currentImg.toDataURL();
  const nextDataUrl = nextImg.toDataURL();

  const currentSlideNum = rectInfo.currentSlide ?? (typeof currentSlide === 'number' ? currentSlide : null);
  const totalSlidesNum = rectInfo.totalSlides ?? null;
  const nextSlideNum = (typeof currentSlideNum === 'number' && typeof totalSlidesNum === 'number')
    ? (currentSlideNum < totalSlidesNum ? currentSlideNum + 1 : null)
    : (typeof currentSlideNum === 'number' ? currentSlideNum + 1 : null);

  return {
    success: true,
    currentSlide: currentSlideNum,
    nextSlide: nextSlideNum,
    totalSlides: totalSlidesNum,
    current: { dataUrl: currentDataUrl },
    next: { dataUrl: nextDataUrl }
  };
}

function getGoogleSession() {
  return session.fromPartition(GOOGLE_SESSION_PARTITION);
}

// Try to ensure Google (Slides/docs) responses are interpreted as UTF-8 when charset is missing.
// Speaker notes can show U+FFFD if responses are decoded with the wrong encoding; this may help.
function setupGoogleSessionEncoding() {
  const googleSession = getGoogleSession();
  googleSession.webRequest.onHeadersReceived({ urls: ['https://docs.google.com/*', 'https://*.google.com/*'] }, (details, callback) => {
    const raw = details.responseHeaders || {};
    const h = {};
    for (const k of Object.keys(raw)) h[k] = Array.isArray(raw[k]) ? raw[k].slice() : [raw[k]];
    const ctKey = Object.keys(h).find((k) => k.toLowerCase() === 'content-type');
    if (ctKey) {
      const ct = h[ctKey];
      const arr = Array.isArray(ct) ? ct : [ct];
      const updated = arr.map((v) => {
        const s = String(v || '');
        if ((s.includes('text/') || s.includes('application/javascript') || s.includes('application/json')) && !/charset\s*=/i.test(s)) {
          return (s.trim().replace(/\s*$/, '') + '; charset=utf-8');
        }
        return v;
      });
      if (updated.some((v, i) => v !== arr[i])) {
        h[ctKey] = updated;
      }
    }
    callback({ responseHeaders: h });
  });
}

// Get preferences file path
function getPreferencesPath() {
  return path.join(app.getPath('userData'), 'preferences.json');
}

// Load preferences
function loadPreferences() {
  try {
    const prefsPath = getPreferencesPath();
    // Intentionally quiet by default: loadPreferences() is called frequently.
    
    if (fs.existsSync(prefsPath)) {
      const data = fs.readFileSync(prefsPath, 'utf8');
      const prefs = JSON.parse(data);
      // Allow preferences to control verbose logging, but never print secrets.
      setVerboseLoggingFromPrefs(prefs);
      // Normalize/migrate preferences in-memory (do not write here; loadPreferences is called often)
      // Primary/Backup migration: support both legacy backupIp1/2/3 and new backupIps[]
      refreshLocalBackupTargetKeySet();
      const mergedBk = mergeBackupIpsFromPrefsObject(prefs);
      const cleanedBk = removeSelfReferentialBackupIps(mergedBk);
      prefs.backupIps = cleanedBk.kept;
      if (cleanedBk.removed.length > 0) {
        logWarn('[Preferences] Stripped backup IP(s) that resolve to this machine:', cleanedBk.removed.join(', '));
        delete prefs.backupIp1;
        delete prefs.backupIp2;
        delete prefs.backupIp3;
        try {
          savePreferences(prefs);
        } catch (e) {
          logError('[Preferences] Could not persist cleaned backup IPs:', e.message);
        }
      }
      prefs.controllerIps = getControllerIpsFromPrefs(prefs);
      prefs.presetUrls = getPresetUrlsFromPrefs(prefs);
      // Removed layout mode 'narrow' (broken preview scaling); treat as Google default.
      if (prefs.notesLayout === 'narrow') prefs.notesLayout = 'default';
      if (!VALID_PRESENTATION_GPU_MODES.has(String(prefs.presentationGpuMode || ''))) {
        prefs.presentationGpuMode = 'default';
      }
      if (prefs.presentationNativeFullscreen !== undefined && prefs.presentationNativeFullscreen !== null) {
        prefs.presentationNativeFullscreen = prefs.presentationNativeFullscreen === true;
      }
      logDebug('[Preferences] Loaded preferences:', safeStringify(prefs));
      return prefs;
    } else {
      logDebug('[Preferences] Preferences file does not exist, returning empty object');
    }
  } catch (error) {
    logError('[Preferences] Error loading preferences:', error);
    logError('[Preferences] Error details:', {
      message: error.message,
      code: error.code,
      path: getPreferencesPath()
    });
  }
  return {};
}

// Save preferences
function savePreferences(prefs) {
  const meta = { removedSelfReferentialBackupIps: [] };
  try {
    const prefsPath = getPreferencesPath();
    // Ensure verbose flag is applied immediately
    setVerboseLoggingFromPrefs(prefs);
    // Normalize/migrate before writing
    prefs.backupIps = getBackupIpsFromPrefs(prefs);
    refreshLocalBackupTargetKeySet();
    const selfBackup = removeSelfReferentialBackupIps(prefs.backupIps);
    meta.removedSelfReferentialBackupIps = selfBackup.removed;
    if (selfBackup.removed.length > 0) {
      logWarn(
        '[Preferences] Removed backup address(es) that resolve to this machine (prevents primary→self loops):',
        selfBackup.removed.join(', ')
      );
    }
    prefs.backupIps = selfBackup.kept;
    prefs.controllerIps = getControllerIpsFromPrefs(prefs);
    if (prefs.defaultNotesZoomSteps !== undefined && prefs.defaultNotesZoomSteps !== null) {
      prefs.defaultNotesZoomSteps = clampNotesZoomSteps(prefs.defaultNotesZoomSteps);
    }
    if (prefs.notesLayout === 'narrow') prefs.notesLayout = 'default';
    logDebug('[Preferences] Saving to:', prefsPath);
    logDebug('[Preferences] Data to save (sanitized):', safeStringify(prefs, 2));
    
    // Ensure directory exists
    const prefsDir = path.dirname(prefsPath);
    if (!fs.existsSync(prefsDir)) {
      logDebug('[Preferences] Creating directory:', prefsDir);
      fs.mkdirSync(prefsDir, { recursive: true });
    }
    
    fs.writeFileSync(prefsPath, JSON.stringify(prefs, null, 2), 'utf8');
    logInfo('[Preferences] Preferences saved');
    invalidateLocalBackupTargetKeyCache();

    // Verify it was written
    if (fs.existsSync(prefsPath)) {
      const stats = fs.statSync(prefsPath);
      logDebug('[Preferences] File verified - size:', stats.size, 'bytes');
    } else {
      logError('[Preferences] ERROR: File was not created after write!');
    }
    return meta;
  } catch (error) {
    logError('[Preferences] Error saving preferences:', error);
    logError('[Preferences] Error details:', {
      message: error.message,
      code: error.code,
      path: getPreferencesPath(),
      stack: error.stack
    });
    throw error; // Re-throw so caller can handle it
  }
}

// Primary/Backup System Functions

// Check if current instance is in backup mode
function isBackupMode() {
  const prefs = loadPreferences();
  return prefs.primaryBackupMode === 'backup';
}

function normalizeBackupIps(ips) {
  if (!Array.isArray(ips)) return [];
  const out = [];
  const seen = new Set();
  for (const raw of ips) {
    const ip = String(raw || '').trim();
    if (!ip) continue;
    if (seen.has(ip)) continue;
    seen.add(ip);
    out.push(ip);
  }
  return out;
}

// Controller allowlist (security)
function normalizeControllerIps(ips) {
  if (!Array.isArray(ips)) return [];
  const out = [];
  const seen = new Set();
  for (const raw of ips) {
    const v = String(raw || '').trim();
    if (!v) continue;
    if (seen.has(v)) continue;
    seen.add(v);
    out.push(v);
  }
  return out;
}

function getControllerIpsFromPrefs(prefs) {
  // Stored in preferences as an array: prefs.controllerIps: string[]
  return normalizeControllerIps(prefs?.controllerIps);
}

function getPresetUrlsFromPrefs(prefs) {
  if (Array.isArray(prefs?.presetUrls)) {
    return prefs.presetUrls.map(u => String(u || '').trim()).filter(Boolean);
  }
  const legacy = [
    prefs?.presentation1,
    prefs?.presentation2,
    prefs?.presentation3
  ].map(u => String(u || '').trim()).filter(Boolean);
  return legacy;
}

function normalizeRemoteAddress(addr) {
  let a = String(addr || '').trim();
  // IPv6-mapped IPv4 (common on Node/Electron)
  if (a.startsWith('::ffff:')) a = a.slice(7);
  // Normalize loopback
  if (a === '::1') a = '127.0.0.1';
  return a;
}

function isLocalhostAddress(addr) {
  const a = normalizeRemoteAddress(addr);
  return a === '127.0.0.1';
}

// --- Never use this machine as its own backup (avoids primary→self HTTP loops on open-presentation, etc.) ---
let cachedLocalBackupTargetKeys = null;
let cachedLocalBackupTargetKeysAt = 0;
const LOCAL_BACKUP_TARGET_CACHE_MS = 15000;

function stripIpv6ZoneIndex(addr) {
  const s = String(addr || '').trim();
  const i = s.indexOf('%');
  if (i === -1) return s;
  const base = s.slice(0, i);
  return net.isIPv6(base) ? base : s;
}

function refreshLocalBackupTargetKeySet() {
  const keys = new Set();
  const addKey = (addr) => {
    const raw = stripIpv6ZoneIndex(addr);
    if (!raw) return;
    const ipType = net.isIP(raw);
    if (ipType === 4) {
      keys.add(raw);
      keys.add(`::ffff:${raw}`);
      return;
    }
    if (ipType === 6) {
      const low = raw.toLowerCase();
      keys.add(low);
      if (low.startsWith('::ffff:')) {
        const v4 = low.slice(7);
        if (net.isIP(v4) === 4) keys.add(v4);
      }
      const asLoop = normalizeRemoteAddress(low);
      if (asLoop === '127.0.0.1') {
        keys.add('127.0.0.1');
        keys.add('::ffff:127.0.0.1');
      }
    }
  };

  addKey('127.0.0.1');
  addKey('::1');

  try {
    const ifaces = os.networkInterfaces();
    for (const name of Object.keys(ifaces || {})) {
      for (const info of ifaces[name] || []) {
        if (!info || !info.address) continue;
        addKey(info.address);
      }
    }
  } catch (e) {
    logDebug('[Backup] networkInterfaces() failed:', e.message);
  }

  cachedLocalBackupTargetKeys = keys;
  cachedLocalBackupTargetKeysAt = Date.now();
  return keys;
}

function getLocalBackupTargetKeySet() {
  if (!cachedLocalBackupTargetKeys || (Date.now() - cachedLocalBackupTargetKeysAt) > LOCAL_BACKUP_TARGET_CACHE_MS) {
    return refreshLocalBackupTargetKeySet();
  }
  return cachedLocalBackupTargetKeys;
}

function invalidateLocalBackupTargetKeyCache() {
  cachedLocalBackupTargetKeys = null;
  cachedLocalBackupTargetKeysAt = 0;
}

function hostMatchesLocalBackupKeys(host, keys) {
  const h = String(host || '').trim();
  if (!h) return false;
  const bare = stripIpv6ZoneIndex(h);
  const ipType = net.isIP(bare);
  if (ipType === 4) {
    return keys.has(bare) || keys.has(`::ffff:${bare}`);
  }
  if (ipType === 6) {
    const low = bare.toLowerCase();
    if (keys.has(low)) return true;
    if (low.startsWith('::ffff:')) {
      const tail = low.slice(7);
      if (net.isIP(tail) === 4) return keys.has(tail) || keys.has(`::ffff:${tail}`);
    }
    if (normalizeRemoteAddress(low) === '127.0.0.1') return keys.has('127.0.0.1');
    return false;
  }
  try {
    const results = dns.lookupSync(h, { all: true });
    for (const rec of results || []) {
      if (rec && rec.address && hostMatchesLocalBackupKeys(rec.address, keys)) return true;
    }
  } catch (e) {
    return false;
  }
  return false;
}

function isSelfReferentialBackupTarget(host) {
  const keys = getLocalBackupTargetKeySet();
  return hostMatchesLocalBackupKeys(host, keys);
}

function removeSelfReferentialBackupIps(ips) {
  const list = normalizeBackupIps(Array.isArray(ips) ? ips : []);
  const kept = [];
  const removed = [];
  for (const ip of list) {
    if (isSelfReferentialBackupTarget(ip)) {
      removed.push(ip);
    } else {
      kept.push(ip);
    }
  }
  return { kept, removed };
}

function parseIpv4ToInt(ip) {
  const parts = String(ip || '').trim().split('.');
  if (parts.length !== 4) return null;
  const nums = parts.map((p) => {
    if (p === '' || !/^\d+$/.test(p)) return null;
    const n = Number(p);
    if (!Number.isInteger(n) || n < 0 || n > 255) return null;
    return n;
  });
  if (nums.some((n) => n === null)) return null;
  // Use unsigned 32-bit
  return (((nums[0] << 24) >>> 0) | (nums[1] << 16) | (nums[2] << 8) | nums[3]) >>> 0;
}

function parseCidr(cidr) {
  const s = String(cidr || '').trim();
  const idx = s.indexOf('/');
  if (idx <= 0) return null;
  const ipStr = s.slice(0, idx).trim();
  const prefixStr = s.slice(idx + 1).trim();
  if (!/^\d+$/.test(prefixStr)) return null;
  const prefix = Number(prefixStr);
  if (!Number.isInteger(prefix) || prefix < 0 || prefix > 32) return null;
  const ipInt = parseIpv4ToInt(ipStr);
  if (ipInt === null) return null;
  const mask = prefix === 0 ? 0 : ((0xffffffff << (32 - prefix)) >>> 0);
  const network = (ipInt & mask) >>> 0;
  return { network, mask, prefix };
}

function isAllowedByControllerEntry(entry, remoteIp) {
  const e = String(entry || '').trim();
  if (!e) return false;
  const r = normalizeRemoteAddress(remoteIp);

  // CIDR entry (IPv4 only)
  if (e.includes('/')) {
    const cidr = parseCidr(e);
    if (!cidr) return false;
    const remoteInt = parseIpv4ToInt(r);
    if (remoteInt === null) return false;
    return ((remoteInt & cidr.mask) >>> 0) === cidr.network;
  }

  // Exact IP match
  return e === r;
}

function isControllerAllowedRequest(req, prefs) {
  // Always allow local requests (desktop app UI and local tooling)
  const remote = normalizeRemoteAddress(req?.socket?.remoteAddress);
  if (isLocalhostAddress(remote)) return true;

  const allowlist = getControllerIpsFromPrefs(prefs);
  if (!allowlist || allowlist.length === 0) return true; // no restrictions

  // Allow if any entry matches (IP or CIDR)
  return allowlist.some((entry) => isAllowedByControllerEntry(entry, remote));
}

/**
 * Web UI served to Cloudflare tunnel clients: cloudflared connects from 127.0.0.1, so we treat
 * localhost + tunnel enabled as "public share" and hide in-browser Settings + block sensitive API proxy routes.
 * LAN clients use real source IPs and get the full UI.
 */
function isWebUiRestrictedTunnelClient(req, prefs) {
  if (!prefs || prefs.cloudflaredEnabled !== true) return false;
  const remote = normalizeRemoteAddress(req?.socket?.remoteAddress);
  return isLocalhostAddress(remote);
}

function getWebUiPinScopeFromPrefs(prefs) {
  const s = String(prefs?.webUiPinScope || 'tunnel').toLowerCase();
  if (s === 'lan' || s === 'both') return s;
  return 'tunnel';
}

/**
 * Whether the Web UI PIN gate should run for this request (PIN must also be configured).
 * tunnel: Cloudflare Quick Tunnel path (localhost to the Web UI while tunnel is on).
 * lan: non-localhost clients only (typical LAN devices). Does not cover the tunnel URL, which appears as localhost.
 * both: tunnel path or non-localhost.
 */
function isWebUiPinGateActiveForRequest(req, prefs) {
  const remote = normalizeRemoteAddress(req?.socket?.remoteAddress);
  const isLocal = isLocalhostAddress(remote);
  const restrictedTunnel = isWebUiRestrictedTunnelClient(req, prefs);
  const scope = getWebUiPinScopeFromPrefs(prefs);
  if (scope === 'lan') return !isLocal;
  if (scope === 'both') return restrictedTunnel || !isLocal;
  return restrictedTunnel;
}

// --- Web UI PIN (optional; scope via webUiPinScope) ---

const TUNNEL_WEB_UI_COOKIE_NAME = 'gso_tunnel_webui';
const WEBUI_TUNNEL_SESSION_SECRET_BYTES = 32;
const TUNNEL_PIN_UNLOCK_MAX_BODY = 4096;
const TUNNEL_PIN_MAX_FAILS = 10;
const TUNNEL_PIN_FAIL_WINDOW_MS = 15 * 60 * 1000;

let webUiTunnelSessionSecret = null;
const tunnelPinFailTracker = new Map();

function sanitizePreferencesForClient(prefs) {
  if (!prefs || typeof prefs !== 'object') return {};
  const out = { ...prefs };
  delete out.webUiTunnelPinScrypt;
  delete out.webUiTunnelPinSalt;
  delete out.webUiTunnelPin;
  delete out.webUiTunnelPinClear;
  out.webUiTunnelPinEnabled = !!(prefs.webUiTunnelPinScrypt && prefs.webUiTunnelPinSalt);
  return out;
}

function getWebUiTunnelSessionSecretPath() {
  return path.join(app.getPath('userData'), 'webui-tunnel-session-secret.bin');
}

function rotateWebUiTunnelSessionSecret() {
  webUiTunnelSessionSecret = crypto.randomBytes(WEBUI_TUNNEL_SESSION_SECRET_BYTES);
  try {
    fs.writeFileSync(getWebUiTunnelSessionSecretPath(), webUiTunnelSessionSecret);
  } catch (e) {
    logError('[Web UI] Failed to write tunnel session secret:', e.message);
  }
}

function loadWebUiTunnelSessionSecret() {
  if (webUiTunnelSessionSecret && webUiTunnelSessionSecret.length) return webUiTunnelSessionSecret;
  const p = getWebUiTunnelSessionSecretPath();
  try {
    if (fs.existsSync(p)) {
      const buf = fs.readFileSync(p);
      if (buf.length >= 16) {
        webUiTunnelSessionSecret = buf;
        return webUiTunnelSessionSecret;
      }
    }
  } catch (e) {
    logError('[Web UI] Failed to read tunnel session secret:', e.message);
  }
  rotateWebUiTunnelSessionSecret();
  return webUiTunnelSessionSecret;
}

function hashWebUiTunnelPin(pin, saltHex) {
  const salt = Buffer.from(String(saltHex || ''), 'hex');
  if (salt.length === 0) throw new Error('invalid salt');
  return crypto.scryptSync(String(pin).normalize('NFKC'), salt, 64, {
    N: 16384,
    r: 8,
    p: 1,
    maxmem: 64 * 1024 * 1024
  }).toString('hex');
}

function isWebUiTunnelPinConfigured(prefs) {
  return !!(prefs && prefs.webUiTunnelPinScrypt && prefs.webUiTunnelPinSalt);
}

function verifyWebUiTunnelPin(pin, prefs) {
  if (!isWebUiTunnelPinConfigured(prefs)) return false;
  if (!/^\d{4,12}$/.test(String(pin || ''))) return false;
  try {
    const got = hashWebUiTunnelPin(pin, prefs.webUiTunnelPinSalt);
    const want = String(prefs.webUiTunnelPinScrypt || '');
    if (got.length !== want.length) return false;
    return crypto.timingSafeEqual(Buffer.from(got, 'hex'), Buffer.from(want, 'hex'));
  } catch (e) {
    return false;
  }
}

function parseCookieHeader(raw) {
  const out = {};
  const s = String(raw || '');
  if (!s) return out;
  for (const part of s.split(';')) {
    const idx = part.indexOf('=');
    if (idx < 0) continue;
    const k = part.slice(0, idx).trim();
    const v = part.slice(idx + 1).trim();
    if (k) out[k] = decodeURIComponent(v);
  }
  return out;
}

function readWebUiTunnelCookie(req) {
  return parseCookieHeader(req.headers && req.headers.cookie)[TUNNEL_WEB_UI_COOKIE_NAME];
}

function tunnelUnlockClientKey(req) {
  const h = req.headers || {};
  const cf = h['cf-connecting-ip'] || h['CF-Connecting-IP'];
  const cfStr = cf ? String(cf).trim() : '';
  if (cfStr) return `cf:${cfStr}`;
  return `ra:${normalizeRemoteAddress(req?.socket?.remoteAddress)}`;
}

function isTunnelPinUnlockBlocked(key) {
  const row = tunnelPinFailTracker.get(key);
  if (!row) return false;
  if (Date.now() > row.resetAt) {
    tunnelPinFailTracker.delete(key);
    return false;
  }
  return row.n >= TUNNEL_PIN_MAX_FAILS;
}

function recordTunnelPinFailure(key) {
  const now = Date.now();
  let row = tunnelPinFailTracker.get(key);
  if (!row || now > row.resetAt) {
    row = { n: 0, resetAt: now + TUNNEL_PIN_FAIL_WINDOW_MS };
  }
  row.n += 1;
  tunnelPinFailTracker.set(key, row);
}

function clearTunnelPinFailures(key) {
  tunnelPinFailTracker.delete(key);
}

function isValidTunnelWebUiSessionCookie(cookieVal, prefs) {
  if (!cookieVal || typeof cookieVal !== 'string') return false;
  if (!isWebUiTunnelPinConfigured(prefs)) return false;
  const parts = cookieVal.split('.');
  if (parts.length !== 2) return false;
  const [payloadB64, sigHex] = parts;
  const secret = loadWebUiTunnelSessionSecret();
  const h = crypto.createHmac('sha256', secret);
  h.update(payloadB64);
  const expect = h.digest('hex');
  if (expect.length !== sigHex.length) return false;
  try {
    if (!crypto.timingSafeEqual(Buffer.from(expect, 'hex'), Buffer.from(sigHex, 'hex'))) return false;
  } catch (e) {
    return false;
  }
  let payload;
  try {
    payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8'));
  } catch (e) {
    return false;
  }
  if (!payload || payload.v !== 1) return false;
  const exp = Number(payload.exp);
  if (!Number.isFinite(exp) || Date.now() > exp) return false;
  const pinTag = String(payload.pt || '');
  const wantTag = String(prefs.webUiTunnelPinScrypt).slice(0, 16);
  if (!wantTag || pinTag !== wantTag) return false;
  return true;
}

function buildTunnelWebUiSessionCookieValue(prefs) {
  const exp = Date.now() + 7 * 24 * 60 * 60 * 1000;
  const pinTag = String(prefs.webUiTunnelPinScrypt).slice(0, 16);
  const payload = Buffer.from(JSON.stringify({ v: 1, exp, pt: pinTag }), 'utf8').toString('base64url');
  const secret = loadWebUiTunnelSessionSecret();
  const h = crypto.createHmac('sha256', secret);
  h.update(payload);
  const sig = h.digest('hex');
  return `${payload}.${sig}`;
}

function shouldUseSecureFlagForTunnelSessionCookie(req) {
  if (webUiServerUsesHttps) return true;
  try {
    const h = req && req.headers;
    if (!h) return false;
    if (h['cf-ray'] || h['CF-Ray']) return true;
    const xf = String(h['x-forwarded-proto'] || '').toLowerCase();
    const first = xf.split(',')[0].trim();
    if (first === 'https') return true;
  } catch (e) {
    return false;
  }
  return false;
}

function buildSetTunnelSessionCookieHeader(prefs, req) {
  const val = buildTunnelWebUiSessionCookieValue(prefs);
  const parts = [
    `${TUNNEL_WEB_UI_COOKIE_NAME}=${encodeURIComponent(val)}`,
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    `Max-Age=${7 * 24 * 60 * 60}`
  ];
  if (shouldUseSecureFlagForTunnelSessionCookie(req)) parts.push('Secure');
  return parts.join('; ');
}

function buildTunnelUnlockHtml() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Unlock Web Remote</title>
  <style>
    * { box-sizing: border-box; }
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; margin: 0; min-height: 100vh; display: flex; align-items: center; justify-content: center; background: #0f172a; color: #e2e8f0; padding: 24px; }
    .card { width: 100%; max-width: 360px; background: #1e293b; border-radius: 12px; padding: 28px 24px; box-shadow: 0 12px 40px rgba(0,0,0,0.35); border: 1px solid #334155; }
    h1 { margin: 0 0 8px; font-size: 1.25rem; font-weight: 600; }
    p { margin: 0 0 20px; font-size: 0.9rem; color: #94a3b8; line-height: 1.45; }
    label { display: block; font-size: 0.8rem; color: #94a3b8; margin-bottom: 6px; }
    input { width: 100%; padding: 12px 14px; border-radius: 8px; border: 1px solid #475569; background: #0f172a; color: #f8fafc; font-size: 1.1rem; letter-spacing: 0.08em; }
    input:focus { outline: 2px solid #38bdf8; border-color: #38bdf8; }
    button { margin-top: 18px; width: 100%; padding: 12px; border: none; border-radius: 8px; background: #38bdf8; color: #0f172a; font-weight: 600; font-size: 1rem; cursor: pointer; }
    button:hover { background: #7dd3fc; }
    button:disabled { opacity: 0.55; cursor: not-allowed; }
    #msg { margin-top: 14px; font-size: 0.85rem; min-height: 1.2em; }
    .err { color: #fca5a5; }
    .ok { color: #86efac; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Web remote locked</h1>
    <p>Enter the PIN configured in the desktop app to use this page.</p>
    <label for="pin">PIN</label>
    <input id="pin" type="password" inputmode="numeric" pattern="[0-9]*" autocomplete="one-time-code" maxlength="12" autofocus />
    <button type="button" id="go">Unlock</button>
    <div id="msg"></div>
  </div>
  <script>
    const pinEl = document.getElementById('pin');
    const goBtn = document.getElementById('go');
    const msg = document.getElementById('msg');
    function setMsg(text, isErr) {
      msg.textContent = text || '';
      msg.className = isErr ? 'err' : (text ? 'ok' : '');
    }
    async function submit() {
      const pin = (pinEl.value || '').trim();
      if (!/^[0-9]{4,12}$/.test(pin)) {
        setMsg('PIN must be 4–12 digits.', true);
        return;
      }
      goBtn.disabled = true;
      setMsg('');
      try {
        const res = await fetch('/tunnel-unlock', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ pin: pin })
        });
        const data = await res.json().catch(function () { return {}; });
        if (res.ok && data.success) {
          setMsg('Success — opening remote…', false);
          window.location.replace('/');
          return;
        }
        if (res.status === 429) {
          setMsg(data.error || 'Too many attempts. Try again later.', true);
        } else {
          setMsg(data.error || 'Incorrect PIN.', true);
        }
      } catch (e) {
        setMsg('Network error. Try again.', true);
      } finally {
        goBtn.disabled = false;
      }
    }
    goBtn.addEventListener('click', submit);
    pinEl.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') submit();
    });
  </script>
</body>
</html>`;
}

function mergeBackupIpsFromPrefsObject(prefs) {
  const fromArray = Array.isArray(prefs?.backupIps) ? prefs.backupIps : [];
  const legacy = [prefs?.backupIp1, prefs?.backupIp2, prefs?.backupIp3];
  return normalizeBackupIps([...fromArray, ...legacy]);
}

function getBackupIpsFromPrefs(prefs) {
  refreshLocalBackupTargetKeySet();
  return removeSelfReferentialBackupIps(mergeBackupIpsFromPrefsObject(prefs)).kept;
}

// Get list of configured backup IP addresses (unlimited, user-configurable)
function getBackupIps() {
  const prefs = loadPreferences();
  if (prefs.primaryBackupMode !== 'primary') return [];
  return getBackupIpsFromPrefs(prefs);
}

/**
 * TCP port on backup machines where the HTTP API is listening.
 * Prefer backupPort; if missing/invalid, use apiPort so primary stays aligned when only API port was changed.
 */
function getOutboundBackupHttpPort(prefs) {
  const p = prefs || loadPreferences();
  const bp = Number(p.backupPort);
  if (Number.isFinite(bp) && bp >= 1 && bp <= 65535) return bp;
  const ap = Number(p.apiPort);
  if (Number.isFinite(ap) && ap >= 1 && ap <= 65535) return ap;
  return DEFAULT_API_PORT;
}

// Send command to all backup machines (fire and forget)
async function sendToBackups(endpoint, data = null) {
  const prefs = loadPreferences();
  if (prefs.primaryBackupMode !== 'primary') {
    return; // Not in primary mode
  }
  if (prefs.backupControlsEnabled === false) {
    return; // Backup forwarding disabled (decoupled mode)
  }

  const backupIps = getBackupIps();
  if (backupIps.length === 0) {
    return; // No backups configured
  }
  
  const port = getOutboundBackupHttpPort(prefs);
  
  logDebug(`[Backup] Broadcasting ${endpoint} to ${backupIps.length} backup(s)`);

  // Send to all backups in parallel (fire and forget - don't wait for responses)
  backupIps.forEach((ip) => {
    if (isSelfReferentialBackupTarget(ip)) {
      logWarn(`[Backup] Skipping self-referential backup target ${ip} for ${endpoint}`);
      return;
    }
    const options = {
      hostname: ip,
      port: port,
      path: endpoint,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      timeout: 2000
    };
    
    const req = http.request(options, (res) => {
      // Success - backup received command
      logDebug(`[Backup] Successfully sent to ${ip}:${port}${endpoint}`);
    });
    
    req.on('error', (err) => {
      // Error - backup didn't receive command (log but don't fail)
      logWarn(`[Backup] Failed to send to ${ip}:${port}${endpoint}:`, err.message);
    });
    
    req.on('timeout', () => {
      req.destroy();
      console.error(`[Backup] Timeout sending to ${ip}:${port}${endpoint}`);
    });
    
    if (data) {
      req.write(JSON.stringify(data));
    }
    req.end();
  });
}

// Check connection status of all backup machines
async function checkBackupStatus() {
  const prefs = loadPreferences();
  if (prefs.primaryBackupMode !== 'primary') {
    return { backups: [] };
  }
  
  const backupIps = getBackupIps();
  if (backupIps.length === 0) {
    return { backups: [] };
  }
  
  const port = getOutboundBackupHttpPort(prefs);
  const backups = backupIps.map(ip => ({ ip, status: 'checking' }));
  
  // Check each backup in parallel
  const promises = backupIps.map((ip, index) => {
    return new Promise((resolve) => {
      const options = {
        hostname: ip,
        port: port,
        path: '/api/status',
        method: 'GET',
        timeout: 2000
      };
      
      const req = http.request(options, (res) => {
        let responseData = '';
        res.on('data', (chunk) => {
          responseData += chunk.toString();
        });
        res.on('end', () => {
          if (res.statusCode === 200) {
            backups[index] = { ip, status: 'connected' };
            resolve();
          } else {
            logDebug(`[Backup] Health check ${ip}:${port} HTTP ${res.statusCode} (expected 200)`);
            backups[index] = { ip, status: 'disconnected' };
            resolve();
          }
        });
      });
      
      req.on('error', (err) => {
        const code = err && err.code ? err.code : 'unknown';
        const msg = err && err.message ? err.message : String(err);
        logDebug(`[Backup] Health check ${ip}:${port} error: ${code} ${msg}`);
        backups[index] = { ip, status: 'disconnected' };
        resolve();
      });
      
      req.on('timeout', () => {
        req.destroy();
        logDebug(`[Backup] Health check ${ip}:${port} timed out`);
        backups[index] = { ip, status: 'disconnected' };
        resolve();
      });
      
      req.end();
    });
  });
  
  await Promise.all(promises);

  return { backups };
}

// Start backup status polling (called when app starts in primary mode)
let backupStatusInterval = null;

function startBackupStatusPolling() {
  stopBackupStatusPolling();
  
  const prefs = loadPreferences();
  if (prefs.primaryBackupMode !== 'primary') {
    return;
  }
  
  // Poll immediately, then every 5 seconds
  checkBackupStatus().catch(err => {
    console.error('[Backup] Error checking backup status:', err);
  });
  
  backupStatusInterval = setInterval(() => {
    checkBackupStatus().catch(err => {
      console.error('[Backup] Error checking backup status:', err);
    });
  }, 5000);
  
  console.log('[Backup] Started backup status polling (5s interval)');
}

function stopBackupStatusPolling() {
  if (backupStatusInterval) {
    clearInterval(backupStatusInterval);
    backupStatusInterval = null;
    console.log('[Backup] Stopped backup status polling');
  }
}

async function reopenPresentationAtSlide(urlToReload, savedSlide, notesWereOpen, savedNotesBounds, savedNotesZoomSteps) {
  notesZoomStepsFromDefault =
    typeof savedNotesZoomSteps === 'number'
      ? clampNotesZoomSteps(savedNotesZoomSteps)
      : getDefaultNotesZoomStepsFromPrefs();

  if (notesWindow && !notesWindow.isDestroyed()) {
    notesWindow.removeAllListeners('closed');
    notesWindow.close();
    notesWindow = null;
  }
  if (presentationWindow && !presentationWindow.isDestroyed()) {
    presentationWindow.removeAllListeners('closed');
    presentationWindow.close();
    presentationWindow = null;
  }
  currentSlide = null;
  await new Promise(resolve => setTimeout(resolve, 200));

  const prefs = loadPreferences();
  const displays = screen.getAllDisplays();
  const presentationDisplayId = Number(prefs.presentationDisplayId);
  const notesDisplayId = Number(prefs.notesDisplayId);
  const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
  const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];

  presentationWindow = new BrowserWindow(getPresentationBrowserWindowOptions(presentationDisplay.bounds));
  if (process.platform === 'darwin') {
    applyPresentationFullscreenChrome(presentationWindow, prefs);
  }
  attachCrashHandlers(presentationWindow, 'presentation');

  presentationWindow.webContents.setWindowOpenHandler(({ url }) => {
    const windowOptions = getSpeakerNotesWindowOptions(notesDisplay);
    return { action: 'allow', overrideBrowserWindowOptions: windowOptions };
  });

  const windowCreatedListener = (event, window) => {
    if (window !== presentationWindow && window !== mainWindow) {
      notesWindow = window;
      onNotesWindowCreated(window);
      attachCrashHandlers(window, 'notes');
      window.webContents.on('before-input-event', (event, input) => {
        if (input.key === 'Escape' && input.type === 'keyDown') {
          event.preventDefault();
          if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
          if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
        }
      });
      const fromReload =
        savedNotesBounds && savedNotesBounds.width > 0 && savedNotesBounds.height > 0
          ? savedNotesBounds
          : null;
      applySpeakerNotesInitialGeometry(window, notesDisplay, fromReload);
      app.removeListener('browser-window-created', windowCreatedListener);
    }
  };
  app.on('browser-window-created', windowCreatedListener);

  presentationWindow.on('closed', () => {
    presentationWindow = null;
    currentSlide = null;
  });
  presentationWindow.webContents.on('before-input-event', (event, input) => {
    if (input.key === 'Escape' && input.type === 'keyDown') {
      event.preventDefault();
      if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
      if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
    }
  });

  const presentUrl = toPresentUrl(urlToReload, savedSlide);
  lastPresentationUrl = urlToReload;
  currentSlide = savedSlide;
  presentationWindow.loadURL(presentUrl);
  presentationWindow.show();
  presentationWindow.focus();
  presentationWindow.once('ready-to-show', () => {
    if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.focus();
  });
  presentationWindow.webContents.once('did-finish-load', async () => {
    if (!presentationWindow || presentationWindow.isDestroyed()) return;
    presentationWindow.focus();
    await new Promise(resolve => setTimeout(resolve, 200));
    if (presentationWindow && !presentationWindow.isDestroyed()) {
      presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
      presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
    }
    // Restore slide position: URL fragment is often ignored; after presentation mode is ready, navigate from slide 1 to savedSlide
    const targetSlide = typeof savedSlide === 'number' && savedSlide > 1 ? savedSlide : 1;
    if (targetSlide > 1) {
      await new Promise(resolve => setTimeout(resolve, 1500));
      if (presentationWindow && !presentationWindow.isDestroyed()) {
        presentationWindow.focus();
        await new Promise(resolve => setTimeout(resolve, 50));
        const count = targetSlide - 1;
        for (let i = 0; i < count; i++) {
          presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Right' });
          presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Right' });
          if (i < count - 1) {
            await new Promise(resolve => setTimeout(resolve, 100));
          }
        }
        currentSlide = targetSlide;
        console.log('[Reload] Restored slide position to', targetSlide);
      }
    }
  });
  await new Promise(resolve => setTimeout(resolve, 2000));
  if (notesWereOpen && presentationWindow && !presentationWindow.isDestroyed()) {
    await new Promise(resolve => setTimeout(resolve, 1000));
    if (presentationWindow && !presentationWindow.isDestroyed()) {
      presentationWindow.focus();
      await new Promise(resolve => setTimeout(resolve, 200));
      presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
      presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
      presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
    }
  }
  if (presentationWindow && !presentationWindow.isDestroyed()) {
    setTimeout(() => presentationWindow.focus(), 500);
  }
}

function attachCrashHandlers(win, label) {
  if (!win || !win.webContents) return;
  win.webContents.on('render-process-gone', (event, details) => {
    const reason = details.reason || 'unknown';
    const exitCode = details.exitCode != null ? details.exitCode : '?';
    const content = [
      `Renderer process gone: ${label}`,
      `Time: ${new Date().toISOString()}`,
      `Reason: ${reason}`,
      `Exit code: ${exitCode}`,
      '',
      '--- Last log lines ---',
      logBuffer.slice(-LOG_TAIL_LINES).join('\n')
    ].join('\n');
    writeCrashReport(`renderer-${label}`, content);
    logError('[Crash] Renderer gone:', label, reason, exitCode);

    if (label === 'presentation') {
      const url = lastPresentationUrl;
      const slide = typeof currentSlide === 'number' ? currentSlide : 1;
      let notesWereOpen = false;
      let savedNotesBounds = null;
      const savedNotesZoomSteps = notesZoomStepsFromDefault;
      if (notesWindow && !notesWindow.isDestroyed()) {
        notesWereOpen = true;
        try { savedNotesBounds = notesWindow.getBounds(); } catch (e) {}
        notesWindow.removeAllListeners('closed');
        notesWindow.close();
        notesWindow = null;
      }
      presentationWindow = null;
      currentSlide = null;
      if (url) {
        dialog.showMessageBox(null, {
          type: 'warning',
          title: 'Presentation Window Closed',
          message: 'The presentation window closed unexpectedly.',
          detail: `Reopen and return to slide ${slide}?`,
          buttons: ['Reopen', 'Cancel'],
          defaultId: 0,
          cancelId: 1
        }).then(({ response }) => {
          if (response === 0) reopenPresentationAtSlide(url, slide, notesWereOpen, savedNotesBounds, savedNotesZoomSteps);
        }).catch(() => {});
      }
    } else if (label === 'notes') {
      let savedNotesBounds = null;
      if (notesWindow && !notesWindow.isDestroyed()) {
        try { savedNotesBounds = notesWindow.getBounds(); } catch (e) {}
        notesWindow = null;
      }
      dialog.showMessageBox(null, {
        type: 'warning',
        title: 'Speaker Notes Closed',
        message: 'The speaker notes window closed unexpectedly.',
        detail: 'Reopen speaker notes?',
        buttons: ['Reopen', 'Cancel'],
        defaultId: 0,
        cancelId: 1
      }).then(({ response }) => {
        if (response === 0 && presentationWindow && !presentationWindow.isDestroyed()) {
          const boundsToRestore = savedNotesBounds;
          const onceCreated = (event, newWin) => {
            if (newWin === presentationWindow || newWin === mainWindow) return;
            app.removeListener('browser-window-created', onceCreated);
            notesWindow = newWin;
            onNotesWindowCreated(newWin);
            attachCrashHandlers(newWin, 'notes');
            if (boundsToRestore && boundsToRestore.width > 0 && boundsToRestore.height > 0) {
              setSpeakerNotesBoundsFromCache(newWin, boundsToRestore);
            }
          };
          app.on('browser-window-created', onceCreated);
          presentationWindow.focus();
          setTimeout(() => {
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
              presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
            }
          }, 100);
        }
      }).catch(() => {});
    }
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 900,
    minWidth: 600,
    minHeight: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    },
    autoHideMenuBar: true,
    resizable: true,
    center: true
  });

  mainWindow.loadFile('index.html');
  attachCrashHandlers(mainWindow, 'main');
  // Open DevTools for main window to see logs
  // mainWindow.webContents.openDevTools();
}

/**
 * Maintainer helper: write README screenshots to docs/images (desktop + Web UI).
 * Run: `yarn capture:readme-screenshots` (or `cross-env GSO_README_CAPTURE=1 electron .`).
 * Temporarily sets Web UI theme to "light" for browser captures, then restores preferences.
 */
async function runReadmeScreenshotCapture() {
  const docsImages = path.join(__dirname, 'docs', 'images');
  if (!fs.existsSync(docsImages)) {
    fs.mkdirSync(docsImages, { recursive: true });
  }

  const savePng = async (win, filename) => {
    if (!win || win.isDestroyed()) return;
    await new Promise((r) => setTimeout(r, 350));
    const image = await win.webContents.capturePage();
    fs.writeFileSync(path.join(docsImages, filename), image.toPNG());
    logInfo('[readme-capture] wrote', path.join('docs', 'images', filename));
  };

  let webWin;
  try {
    if (webUiServerUsesHttps) {
      session.defaultSession.setCertificateVerifyProc((request, callback) => {
        if (request.hostname === '127.0.0.1' || request.hostname === 'localhost') {
          callback(0);
        } else {
          callback(-2);
        }
      });
    }

    if (mainWindow && !mainWindow.isDestroyed()) {
      await mainWindow.webContents.executeJavaScript(`
        (function () {
          document.querySelector('[data-target="dashboard"]')?.click();
        })();
      `);
      await new Promise((r) => setTimeout(r, 700));
      await savePng(mainWindow, 'desktop-settings-dashboard.png');

      await mainWindow.webContents.executeJavaScript(`
        (function () {
          document.querySelector('[data-target="presets"]')?.click();
        })();
      `);
      await new Promise((r) => setTimeout(r, 700));
      await savePng(mainWindow, 'desktop-settings-presets.png');

      await mainWindow.webContents.executeJavaScript(`
        (function () {
          document.querySelector('[data-target="advanced"]')?.click();
        })();
      `);
      await new Promise((r) => setTimeout(r, 500));
      await mainWindow.webContents.executeJavaScript(`
        (function () {
          const titles = [...document.querySelectorAll('.panel-title')];
          const p = titles.find((el) => (el.textContent || '').includes('Primary'));
          if (p) p.scrollIntoView({ block: 'start' });
        })();
      `);
      await new Promise((r) => setTimeout(r, 500));
      await savePng(mainWindow, 'desktop-settings-primary-backup.png');
    }

    const port = currentWebUiPort;
    if (!port) {
      logWarn('[readme-capture] Web UI port not ready; skipping in-browser captures');
      return;
    }

    savePreferences({ ...loadPreferences(), webUiTheme: 'light' });

    webWin = new BrowserWindow({
      width: 440,
      height: 920,
      show: false,
      backgroundColor: '#fbfaf7',
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
      },
    });

    const origin = `${webUiServerUsesHttps ? 'https' : 'http'}://127.0.0.1:${port}`;
    await webWin.loadURL(`${origin}/?readme-capture=${Date.now()}`);
    await webWin.webContents.executeJavaScript('document.fonts.ready');
    await new Promise((r) => setTimeout(r, 2000));
    await savePng(webWin, 'web-ui-remote-light.png');

    await webWin.webContents.executeJavaScript(`
      (function () {
        const b = document.querySelector('.bottom-tabs button[data-tab="controls"]') ||
          document.querySelector('.tabs .tab-btn[data-tab="controls"]');
        b?.click();
      })();
    `);
    await new Promise((r) => setTimeout(r, 900));
    await savePng(webWin, 'web-ui-controls-light.png');
  } catch (err) {
    logError('[readme-capture] failed:', err);
  } finally {
    try {
      if (readmeCapturePrefsBackup != null) {
        savePreferences(JSON.parse(JSON.stringify(readmeCapturePrefsBackup)));
      }
    } catch (e) {
      logError('[readme-capture] restore prefs failed:', e);
    }
    if (webWin && !webWin.isDestroyed()) {
      webWin.close();
    }
    if (webUiServerUsesHttps) {
      try {
        session.defaultSession.setCertificateVerifyProc(null);
      } catch (e) {
        /* ignore */
      }
    }
  }
}

function scheduleReadmeScreenshotCapture() {
  if (process.env.GSO_README_CAPTURE !== '1') return;
  const delayMs = Number(process.env.GSO_README_CAPTURE_DELAY_MS) || 5000;
  setTimeout(() => {
    runReadmeScreenshotCapture()
      .catch((err) => logError('[readme-capture]', err))
      .finally(() => {
        setTimeout(() => app.exit(0), 250);
      });
  }, delayMs);
}

// Get all available displays
ipcMain.handle('get-displays', async () => {
  const displays = screen.getAllDisplays();
  return displays.map((display, index) => ({
    id: display.id,
    label: `Monitor ${index + 1} (${display.bounds.width}x${display.bounds.height})`,
    bounds: display.bounds,
    primary: display.bounds.x === 0 && display.bounds.y === 0
  }));
});

// Get saved preferences
ipcMain.handle('get-preferences', async () => {
  return sanitizePreferencesForClient(loadPreferences());
});

// Get speaker notes (normalized) for desktop UI - same as GET /api/get-speaker-notes
ipcMain.handle('get-speaker-notes', async () => {
  if (!notesWindow || notesWindow.isDestroyed()) {
    return { success: false, notes: '', error: 'No speaker notes window is open' };
  }
  try {
    const notesContent = await notesWindow.webContents.executeJavaScript(`
      (function(){
        var el = document.querySelector('div.punch-viewer-speakernotes-text-body-scrollable');
        if (!el) return 'No notes available for this slide.';
        var raw = (el.innerText || el.textContent || '').trim();
        return raw.length > 0 ? raw : 'No notes available for this slide.';
      })()
    `);
    const rawNotes = notesContent || 'No notes available for this slide.';
    const replacementCharsFound = countReplacementChars(rawNotes);
    const encodingIssuesDetected = replacementCharsFound > 0;
    lastNotesEncodingIssue = encodingIssuesDetected;
    const notes = normalizeSpeakerNotes(rawNotes) || 'No notes available for this slide.';
    return { success: true, notes, encodingIssuesDetected, replacementCharsFound };
  } catch (error) {
    console.error('[IPC] Error getting speaker notes:', error);
    return { success: false, notes: '', error: error.message };
  }
});

// Crash report paths for Debug Logs UI
ipcMain.handle('get-crash-info', async () => {
  try {
    const crashReportsDir = getCrashReportsDir();
    let crashDumpsPath = '';
    try {
      crashDumpsPath = app.getPath('crashDumps');
    } catch (e) {
      crashDumpsPath = '(not set)';
    }
    return {
      crashReportsDir,
      crashDumpsPath,
      lastCrashPath: lastCrashPath || '',
      lastCrashTime: lastCrashTime || ''
    };
  } catch (e) {
    return { crashReportsDir: '', crashDumpsPath: '', lastCrashPath: '', lastCrashTime: '', error: e.message };
  }
});

ipcMain.handle('open-crash-reports-folder', async () => {
  try {
    const dir = getCrashReportsDir();
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    await shell.openPath(dir);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get build info (version and build number)
ipcMain.handle('get-build-info', async () => {
  try {
    return {
      version: appBuildInfo.version,
      buildNumber: appBuildInfo.buildNumber,
      chrome: process.versions.chrome,
      electronRuntime: process.versions.electron,
      node: process.versions.node
    };
  } catch (error) {
    console.error('[Build Info] Error loading build info:', error);
    return {
      version: 'unknown',
      buildNumber: 'unknown',
      chrome: process.versions.chrome,
      electronRuntime: process.versions.electron,
      node: process.versions.node
    };
  }
});

// Get network interfaces and IP addresses
ipcMain.handle('get-network-info', async () => {
  const interfaces = os.networkInterfaces();
  const ipAddresses = [];
  
  // Get all IPv4 addresses (excluding internal/loopback, but including localhost)
  Object.keys(interfaces).forEach((ifaceName) => {
    interfaces[ifaceName].forEach((iface) => {
      // Include IPv4 addresses (both internal and external)
      if (iface.family === 'IPv4') {
        ipAddresses.push({
          address: iface.address,
          internal: iface.internal,
          interface: ifaceName
        });
      }
    });
  });
  
  // Sort: non-internal first, then by interface name
  ipAddresses.sort((a, b) => {
    if (a.internal !== b.internal) {
      return a.internal ? 1 : -1;
    }
    return a.interface.localeCompare(b.interface);
  });
  
  return ipAddresses;
});

// Cloudflare Quick Tunnel (desktop Settings only)
ipcMain.handle('get-tunnel-status', () => ({
  enabled: !!loadPreferences().cloudflaredEnabled,
  url: tunnelUrl,
  running: !!cloudflaredProcess
}));

ipcMain.handle('set-tunnel-enabled', async (_event, enabled) => {
  const prefs = loadPreferences();
  prefs.cloudflaredEnabled = !!enabled;
  savePreferences(prefs);
  if (enabled) {
    if (!cloudflaredProcess) {
      startCloudflaredTunnel();
    }
  } else {
    stopCloudflaredTunnel();
  }
  return { success: true };
});

ipcMain.handle('open-external', async (_event, url) => {
  if (typeof url === 'string' && (url.startsWith('http://') || url.startsWith('https://'))) {
    await shell.openExternal(url);
  }
});

// Save preferences
ipcMain.handle('save-preferences', async (event, incoming) => {
  const currentPrefs = loadPreferences();
  const mergedPrefs = { ...currentPrefs, ...incoming };
  delete mergedPrefs.webUiTunnelPin;
  // Never accept PIN hash/salt from IPC; only set via clear / new PIN below.
  delete mergedPrefs.webUiTunnelPinScrypt;
  delete mergedPrefs.webUiTunnelPinSalt;
  if (incoming && incoming.webUiTunnelPinClear === true) {
    rotateWebUiTunnelSessionSecret();
  } else if (incoming && typeof incoming.webUiTunnelPin === 'string' && incoming.webUiTunnelPin.trim() !== '') {
    const salt = crypto.randomBytes(16).toString('hex');
    mergedPrefs.webUiTunnelPinSalt = salt;
    mergedPrefs.webUiTunnelPinScrypt = hashWebUiTunnelPin(incoming.webUiTunnelPin.trim(), salt);
    rotateWebUiTunnelSessionSecret();
  } else {
    mergedPrefs.webUiTunnelPinScrypt = currentPrefs.webUiTunnelPinScrypt;
    mergedPrefs.webUiTunnelPinSalt = currentPrefs.webUiTunnelPinSalt;
  }
  delete mergedPrefs.webUiTunnelPinClear;
  if (incoming && incoming.webUiPinScope !== undefined) {
    const ps = String(incoming.webUiPinScope).toLowerCase();
    mergedPrefs.webUiPinScope = ps === 'lan' || ps === 'both' ? ps : 'tunnel';
  }
  if (incoming && incoming.presentationGpuMode !== undefined) {
    const g = String(incoming.presentationGpuMode || '');
    mergedPrefs.presentationGpuMode = VALID_PRESENTATION_GPU_MODES.has(g) ? g : 'default';
  }
  if (incoming && incoming.presentationNativeFullscreen !== undefined) {
    mergedPrefs.presentationNativeFullscreen = incoming.presentationNativeFullscreen === true;
  }
  const saveMeta = savePreferences(mergedPrefs) || {};
  return { success: true, ...saveMeta };
});

ipcMain.handle('relaunch-speaker-notes', async () => {
  try {
    return await relaunchSpeakerNotesWindow();
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Open file dialog for selecting a custom CSS file (Web UI white-label)
ipcMain.handle('show-open-css-dialog', async () => {
  const win = BrowserWindow.getFocusedWindow();
  const result = await dialog.showOpenDialog(win || BrowserWindow.getAllWindows()[0], {
    title: 'Select custom CSS file',
    properties: ['openFile'],
    filters: [{ name: 'CSS', extensions: ['css'] }, { name: 'All Files', extensions: ['*'] }]
  });
  if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
    return { canceled: true, filePath: null };
  }
  return { canceled: false, filePath: result.filePaths[0] };
});

ipcMain.handle('show-open-logo-dialog', async () => {
  const win = BrowserWindow.getFocusedWindow();
  const result = await dialog.showOpenDialog(win || BrowserWindow.getAllWindows()[0], {
    title: 'Select brand logo image',
    properties: ['openFile'],
    filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'svg'] }, { name: 'All Files', extensions: ['*'] }]
  });
  if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
    return { canceled: true, filePath: null };
  }
  return { canceled: false, filePath: result.filePaths[0] };
});

ipcMain.handle('show-open-cert-dialog', async () => {
  const win = BrowserWindow.getFocusedWindow();
  const result = await dialog.showOpenDialog(win || BrowserWindow.getAllWindows()[0], {
    title: 'Select TLS certificate file',
    properties: ['openFile'],
    filters: [{ name: 'Certificate', extensions: ['pem', 'crt', 'cer'] }, { name: 'All Files', extensions: ['*'] }]
  });
  if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
    return { canceled: true, filePath: null };
  }
  return { canceled: false, filePath: result.filePaths[0] };
});

ipcMain.handle('show-open-key-dialog', async () => {
  const win = BrowserWindow.getFocusedWindow();
  const result = await dialog.showOpenDialog(win || BrowserWindow.getAllWindows()[0], {
    title: 'Select TLS private key file',
    properties: ['openFile'],
    filters: [{ name: 'Key', extensions: ['pem', 'key'] }, { name: 'All Files', extensions: ['*'] }]
  });
  if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
    return { canceled: true, filePath: null };
  }
  return { canceled: false, filePath: result.filePaths[0] };
});

// Template CSS for Web UI white-label (colors only, no layout)
const WEB_UI_CSS_TEMPLATE = `/*
  Web UI Custom Style Template
  Use this file as a starting point to white-label the Web UI.
  Edit the variables or rules below, then select this file (or your copy) as "Custom CSS file" in the app.
  This template only overrides colors; layout is controlled by the app.
*/

:root {
  /* Primary (buttons, links, focus) */
  --gso-primary: #667eea;
  --gso-primary-hover: #5568d3;
  --gso-primary-end: #764ba2;

  /* Backgrounds */
  --gso-bg-body-start: #667eea;
  --gso-bg-body-end: #764ba2;
  --gso-bg-container: #ffffff;
  --gso-bg-panels: #f8f9fa;

  /* Text */
  --gso-text: #333333;
  --gso-text-muted: #666666;

  /* Borders */
  --gso-border: #e0e0e0;

  /* Secondary (secondary buttons) */
  --gso-secondary: #6c757d;
  --gso-secondary-hover: #5a6268;
}

/* Body gradient */
body {
  background: linear-gradient(135deg, var(--gso-bg-body-start) 0%, var(--gso-bg-body-end) 100%) !important;
}

/* Main card */
.container {
  background: var(--gso-bg-container) !important;
}

/* Headings */
h1, h2, h3 {
  color: var(--gso-text) !important;
}

/* System icon (when no custom logo) */
.system-icon {
  color: var(--gso-primary) !important;
}

/* Primary buttons */
.btn:not(.btn-secondary),
.remote-btn-prev,
.remote-btn-next,
.notes-toggle-btn.active,
.preview-toggle-btn.active {
  background: var(--gso-primary) !important;
  color: white !important;
}
.btn:not(.btn-secondary):hover,
.remote-btn-prev:hover,
.remote-btn-next:hover {
  background: var(--gso-primary-hover) !important;
}

/* Secondary buttons */
.btn-secondary,
.notes-toggle-btn:not(.active),
.preview-toggle-btn:not(.active) {
  background: var(--gso-secondary) !important;
}
.btn-secondary:hover {
  background: var(--gso-secondary-hover) !important;
}

/* Tabs */
.tab-btn.active {
  color: var(--gso-primary) !important;
  border-bottom-color: var(--gso-primary) !important;
}

/* Inputs */
input[type="text"] {
  border-color: var(--gso-border) !important;
}
input[type="text"]:focus {
  border-color: var(--gso-primary) !important;
}

/* Panels (notes, previews) */
.slide-previews-grid,
.speaker-notes-content-wrapper {
  background: var(--gso-bg-panels) !important;
  border-color: var(--gso-border) !important;
}

/* Stagetimer (default gradient) */
.stagetimer-container:not(.disabled):not(.error) {
  background: linear-gradient(135deg, var(--gso-primary) 0%, var(--gso-primary-end) 100%) !important;
}
`;

ipcMain.handle('download-css-template', async () => {
  try {
    const win = BrowserWindow.getFocusedWindow();
    const defaultName = 'gslide-opener-web-ui-template.css';
    const result = await dialog.showSaveDialog(win || BrowserWindow.getAllWindows()[0], {
      title: 'Save CSS template',
      defaultPath: path.join(app.getPath('downloads'), defaultName),
      filters: [{ name: 'CSS', extensions: ['css'] }, { name: 'All Files', extensions: ['*'] }]
    });
    if (result.canceled || !result.filePath) {
      return { success: false, canceled: true };
    }
    fs.writeFileSync(result.filePath, WEB_UI_CSS_TEMPLATE, 'utf8');
    return { success: true, filePath: result.filePath };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Desktop debug log access
ipcMain.handle('get-log-buffer', async () => {
  return { lines: logBuffer.slice() };
});

ipcMain.handle('clear-log-buffer', async () => {
  logBuffer = [];
  return { success: true };
});

ipcMain.handle('export-log-buffer', async () => {
  try {
    const prefs = loadPreferences();
    const header = [
      'Google Slides Opener - Debug Logs',
      `Generated: ${new Date().toISOString()}`,
      `Version: v${appBuildInfo.version}.${appBuildInfo.buildNumber}`,
      `Platform: ${process.platform}`,
      '',
      '--- Preferences (sanitized) ---',
      safeStringify(prefs, 2),
      '',
      '--- Log output ---',
      ''
    ].join('\n');

    const content = header + logBuffer.join('\n') + '\n';

    const defaultName = `gslide-opener-debug-${new Date().toISOString().replace(/[:.]/g, '-')}.txt`;
    const result = await dialog.showSaveDialog({
      title: 'Save Debug Log',
      defaultPath: path.join(app.getPath('downloads'), defaultName),
      filters: [{ name: 'Text Files', extensions: ['txt'] }]
    });

    if (result.canceled || !result.filePath) {
      return { success: false, canceled: true };
    }

    fs.writeFileSync(result.filePath, content, 'utf8');
    return { success: true, filePath: result.filePath };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Sign in with Google
ipcMain.handle('google-signin', async () => {
  const googleSession = getGoogleSession();
  
  const authWindow = new BrowserWindow({
    width: 500,
    height: 600,
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      partition: GOOGLE_SESSION_PARTITION
    }
  });

  // Navigate to Google Sign In
  authWindow.loadURL('https://accounts.google.com/signin');

  authWindow.once('ready-to-show', () => {
    authWindow.show();
  });

  // Listen for successful authentication
  return new Promise((resolve, reject) => {
    let resolved = false;
    
    authWindow.webContents.on('did-navigate', (event, url) => {
      // Check if we've successfully signed in (redirected to myaccount or other Google service)
      if (url.includes('myaccount.google.com') || url.includes('accounts.google.com/ServiceLogin')) {
        if (!resolved) {
          resolved = true;
          authWindow.close();
          resolve({ success: true, message: 'Successfully signed in to Google' });
        }
      }
    });

    authWindow.on('closed', () => {
      if (!resolved) {
        resolved = true;
        reject({ success: false, message: 'Authentication window closed' });
      }
    });
  });
});

// Check if user is already signed in
ipcMain.handle('check-signin-status', async () => {
  try {
    const googleSession = getGoogleSession();
    const cookies = await googleSession.cookies.get({ domain: '.google.com' });
    
    // Check if we have Google authentication cookies
    const hasAuthCookies = cookies.some(cookie => 
      cookie.name === 'SID' || cookie.name === 'HSID' || cookie.name === 'SSID'
    );
    
    let userEmail = null;
    let userName = null;
    
    if (hasAuthCookies) {
      // Try to get user email from cookies
      const emailCookie = cookies.find(cookie => 
        cookie.name === 'Email' || cookie.name === 'email' || cookie.domain.includes('google.com')
      );
      if (emailCookie && emailCookie.value && emailCookie.value.includes('@')) {
        userEmail = emailCookie.value;
      }
      
      // Try to get user name from cookies
      const nameCookie = cookies.find(cookie => 
        cookie.name === 'Name' || cookie.name === 'name'
      );
      if (nameCookie) {
        userName = nameCookie.value;
      }
      
      // If we don't have email, try to fetch from Google account page
      if (!userEmail) {
        try {
          const tempWindow = new BrowserWindow({
            show: false,
            webPreferences: {
              nodeIntegration: false,
              contextIsolation: true,
              partition: GOOGLE_SESSION_PARTITION
            }
          });
          
          await tempWindow.loadURL('https://myaccount.google.com/');
          await new Promise(resolve => setTimeout(resolve, 2000));
          
          const userInfo = await tempWindow.webContents.executeJavaScript(`
            (function() {
              try {
                var email = null;
                var name = null;
                
                // Look for email in page
                var emailEl = document.querySelector('[data-email]') || 
                             document.querySelector('input[type="email"][value]');
                if (emailEl) {
                  email = emailEl.getAttribute('data-email') || emailEl.value;
                }
                
                // Look for name
                var nameEl = document.querySelector('[data-name]') ||
                            document.querySelector('h1');
                if (nameEl) {
                  name = nameEl.getAttribute('data-name') || nameEl.textContent.trim();
                }
                
                // Try to extract from page title
                if (!email) {
                  var title = document.title;
                  var emailMatch = title.match(/([\\w.-]+@[\\w.-]+\\.[\\w.-]+)/);
                  if (emailMatch) email = emailMatch[1];
                }
                
                return { email: email || null, name: name || null };
              } catch (e) {
                return { email: null, name: null };
              }
            })()
          `);
          
          if (userInfo.email) userEmail = userInfo.email;
          if (userInfo.name) userName = userInfo.name;
          
          tempWindow.close();
        } catch (error) {
          console.error('Error fetching user info:', error);
        }
      }
    }
    
    return { 
      signedIn: hasAuthCookies,
      userEmail: userEmail || null,
      userName: userName || null
    };
  } catch (error) {
    console.error('Error checking sign-in status:', error);
    return { signedIn: false, userEmail: null, userName: null };
  }
});

// Sign out from Google
ipcMain.handle('google-signout', async () => {
  const googleSession = getGoogleSession();
  
  // Clear all cookies and storage data for the Google session
  await googleSession.clearStorageData({
    storages: ['cookies', 'localstorage', 'sessionstorage', 'cachestorage']
  });
  
  return { success: true, message: 'Successfully signed out' };
});

// Open test presentation (present URL opens directly in presentation/slideshow mode)
ipcMain.handle('open-test-presentation', async () => {
  const testUrl = 'https://docs.google.com/presentation/d/1qKhywpFhjG4tAtA1e2Rk9dB2lVk_uu5_Ol5TaBhvvPo/present';
  
  // Load preferences to get selected displays
  const prefs = loadPreferences();
  logDebug('[Test] Loaded preferences:', safeStringify(prefs));
  
  const displays = screen.getAllDisplays();
  logDebug('[Test] All available displays:');
  displays.forEach((display, index) => {
    logDebug(`  Display ${index + 1} - ID: ${display.id}, Bounds: ${safeStringify(display.bounds)}`);
  });
  
  // Convert IDs to numbers for comparison (they might be saved as strings)
  const presentationDisplayId = Number(prefs.presentationDisplayId);
  const notesDisplayId = Number(prefs.notesDisplayId);
  
  const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
  const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];
  
  logDebug('[Test] Selected presentation display ID:', prefs.presentationDisplayId, '(converted to:', presentationDisplayId, ')');
  logDebug('[Test] Resolved presentation display:', presentationDisplay.id, 'Bounds:', presentationDisplay.bounds);
  logDebug('[Test] Selected notes display ID:', prefs.notesDisplayId, '(converted to:', notesDisplayId, ')');
  logDebug('[Test] Resolved notes display:', notesDisplay.id, 'Bounds:', notesDisplay.bounds);
  
  if (!presentationWindow) {
    // Note: Don't use fullscreen: true in constructor as it creates a new Space on macOS
    // We'll use setSimpleFullScreen() after creation to avoid Spaces conflicts
    presentationWindow = new BrowserWindow(getPresentationBrowserWindowOptions(presentationDisplay.bounds));
    
    // Set simple fullscreen on macOS to avoid Spaces conflicts
    if (process.platform === 'darwin') {
      applyPresentationFullscreenChrome(presentationWindow, prefs);
    }
    attachCrashHandlers(presentationWindow, 'presentation');

    presentationWindow.on('closed', () => {
      presentationWindow = null;
      currentSlide = null;
    });
    
    // Listen for Escape key to close both windows
    presentationWindow.webContents.on('before-input-event', (event, input) => {
      if (input.key === 'Escape' && input.type === 'keyDown') {
        logDebug('[Test] Escape pressed, closing presentation and notes windows');
        event.preventDefault(); // Prevent Google Slides from handling Escape
        if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
        if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
      }
    });

    // Handle the speaker notes popup window (open at notes-display size so Slides uses narrow-preview layout)
    presentationWindow.webContents.setWindowOpenHandler((details) => {
      logDebug('[Test] Window open intercepted:', details.url);
      const windowOptions = getSpeakerNotesWindowOptions(notesDisplay);
      return { action: 'allow', overrideBrowserWindowOptions: windowOptions };
    });
    
    // Listen for new windows being created (this will be the notes window)
    const testWindowCreatedListener = (event, window) => {
      if (window !== presentationWindow && window !== mainWindow) {
        logDebug('[Test] Notes window created');
        logDebug('[Test] Presentation display ID:', presentationDisplay.id);
        logDebug('[Test] Notes display ID:', notesDisplay.id);
        notesWindow = window;
        onNotesWindowCreated(window);
        attachCrashHandlers(window, 'notes');
        const initialBounds = window.getBounds();
        logDebug('[Test] Initial window bounds:', initialBounds);
        
        // Add Escape key handler to notes window as well
        window.webContents.on('before-input-event', (event, input) => {
          if (input.key === 'Escape' && input.type === 'keyDown') {
            logDebug('[Test] Escape pressed in notes window, closing all windows');
            event.preventDefault();
            if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
            if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
          }
        });
        
        applySpeakerNotesInitialGeometry(window, notesDisplay, null);

        app.removeListener('browser-window-created', testWindowCreatedListener);
      }
    };
    app.on('browser-window-created', testWindowCreatedListener);
  }

  lastPresentationUrl = testUrl; // Store for reload
  currentSlide = 1;
  resetNotesZoomForNewPresentation();
  presentationWindow.loadURL(testUrl);
  presentationWindow.show();
  
  logDebug('[Test] Window opened, loading URL...');
  
  // Set up navigation listener
  const navigationListener = async (event, url) => {
    logDebug('[Test] Navigated to:', url);
    
    // Just log navigation, don't auto-launch notes
    logDebug('[Test] Navigated to:', url);
  };
  
  presentationWindow.webContents.on('did-navigate', navigationListener);
  
  // Test URL uses /present so it already opens in slideshow mode; no need to send Ctrl+Shift+F5
  presentationWindow.webContents.once('did-finish-load', () => {
    if (!presentationWindow || presentationWindow.isDestroyed()) return;
    logDebug('[Test] Page loaded (presentation mode):', presentationWindow.webContents.getURL());
  });
  
  return { success: true };
});

// Open presentation on specific monitor
ipcMain.handle('open-presentation', async (event, { url, presentationDisplayId, notesDisplayId }) => {
  const displays = screen.getAllDisplays();
  logDebug('[Multi-Monitor] All available displays:');
  displays.forEach((display, index) => {
    logDebug(`  Display ${index + 1} - ID: ${display.id}, Bounds: ${safeStringify(display.bounds)}`);
  });
  
  // Convert IDs to numbers for comparison (they might be passed as strings)
  const presentationDisplayIdNum = Number(presentationDisplayId);
  const notesDisplayIdNum = Number(notesDisplayId);
  
  const presentationDisplay = displays.find(d => d.id === presentationDisplayIdNum);
  // Match reopen/API paths: stale or missing notes ID should still map to a real display, not only primary.
  const notesDisplay = displays.find(d => d.id === notesDisplayIdNum) || displays[0];

  logDebug('[Multi-Monitor] Selected presentation display ID:', presentationDisplayId, '(converted to:', presentationDisplayIdNum, ')');
  logDebug('[Multi-Monitor] Resolved presentation display:', presentationDisplay ? presentationDisplay.id : 'NOT FOUND', 'Bounds:', presentationDisplay ? presentationDisplay.bounds : 'N/A');
  logDebug('[Multi-Monitor] Selected notes display ID:', notesDisplayId, '(converted to:', notesDisplayIdNum, ')');
  logDebug('[Multi-Monitor] Resolved notes display:', notesDisplay ? notesDisplay.id : 'NOT FOUND', 'Bounds:', notesDisplay ? notesDisplay.bounds : 'N/A');

  if (!presentationDisplay) {
    return { success: false, message: 'Invalid presentation display' };
  }

  // Close existing windows if any
  if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
  if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
  currentSlide = null;

  // Open presentation window
  // Note: Don't use fullscreen: true in constructor as it creates a new Space on macOS
  // We'll use setSimpleFullScreen() after creation to avoid Spaces conflicts
  presentationWindow = new BrowserWindow(getPresentationBrowserWindowOptions(presentationDisplay.bounds));
  
  // Set simple fullscreen on macOS to avoid Spaces conflicts
  if (process.platform === 'darwin') {
    applyPresentationFullscreenChrome(presentationWindow, loadPreferences());
  }
  attachCrashHandlers(presentationWindow, 'presentation');

  // Handle the speaker notes popup window (open at notes-display size so Slides uses narrow-preview layout)
  presentationWindow.webContents.setWindowOpenHandler((details) => {
    logDebug('[Multi-Monitor] Window open intercepted:', details.url);
    const windowOptions = getSpeakerNotesWindowOptions(notesDisplay);
    return { action: 'allow', overrideBrowserWindowOptions: windowOptions };
  });
  
  // Listen for new windows being created (this will be the notes window)
  const windowCreatedListener = (event, window) => {
    // Check if this is not the presentation window or main window
    if (window !== presentationWindow && window !== mainWindow) {
      logDebug('[Multi-Monitor] Notes window created');
      logDebug('[Multi-Monitor] Presentation display ID:', presentationDisplayIdNum);
      logDebug('[Multi-Monitor] Notes display ID:', notesDisplayIdNum);
      logDebug('[Multi-Monitor] Notes display object:', notesDisplay);
      notesWindow = window;
      onNotesWindowCreated(window);
      attachCrashHandlers(window, 'notes');
      // Get initial window bounds
      const initialBounds = window.getBounds();
      logDebug('[Multi-Monitor] Initial window bounds:', initialBounds);
      
      // Add Escape key handler to notes window as well
      window.webContents.on('before-input-event', (event, input) => {
        if (input.key === 'Escape' && input.type === 'keyDown') {
          logDebug('[Multi-Monitor] Escape pressed in notes window, closing all windows');
          event.preventDefault();
          if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
          if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
        }
      });
      
      applySpeakerNotesInitialGeometry(window, notesDisplay, null);

      // Remove listener after notes window is created
      app.removeListener('browser-window-created', windowCreatedListener);
    }
  };
  app.on('browser-window-created', windowCreatedListener);

  // Load presentation URL
  lastPresentationUrl = url; // Store for reload
  currentSlide = 1;
  resetNotesZoomForNewPresentation();
  presentationWindow.loadURL(url);

  logDebug('[Multi-Monitor] Window opened, loading URL...');

  // Listen for all page loads
  // Set up navigation listener to detect presentation mode activation
  const navigationListener = async (event, url) => {
    logDebug('[Multi-Monitor] Navigated to:', url);
    // Just log navigation, don't auto-launch notes
  };
  
  presentationWindow.webContents.on('did-navigate', navigationListener);
  
  // Listen for page load, then immediately trigger presentation mode
  presentationWindow.webContents.once('did-finish-load', async () => {
    if (!presentationWindow || presentationWindow.isDestroyed()) return;
    
    const currentUrl = presentationWindow.webContents.getURL();
    logDebug('[Multi-Monitor] Page loaded:', currentUrl);
    
    // Small delay to ensure page is fully interactive
    await new Promise(resolve => setTimeout(resolve, 200));
    
    if (!presentationWindow || presentationWindow.isDestroyed()) return;
    
    logDebug('[Multi-Monitor] Triggering Ctrl+Shift+F5 to enter presentation mode...');
    
    try {
      // Focus the window first to ensure it receives the keyboard events
      presentationWindow.focus();
      await new Promise(resolve => setTimeout(resolve, 50));
      
      // Send real keyboard input events
      presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
      presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
      
      logDebug('[Multi-Monitor] Ctrl+Shift+F5 sent via sendInputEvent');
    } catch (error) {
      logError('[Multi-Monitor] Error sending Ctrl+Shift+F5:', error);
    }
    
    // No auto-launch of speaker notes - user must call open-speaker-notes separately
  });

  presentationWindow.on('closed', () => {
    presentationWindow = null;
    currentSlide = null;
  });
  
  // Listen for Escape key to close both windows
  presentationWindow.webContents.on('before-input-event', (event, input) => {
    if (input.key === 'Escape' && input.type === 'keyDown') {
      logDebug('[Multi-Monitor] Escape pressed, closing presentation and notes windows');
      event.preventDefault(); // Prevent Google Slides from handling Escape
      if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
      if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
    }
  });

  return { success: true };
});

// HTTP API for Bitfocus Companion integration
// Ports are configurable via preferences, defaults below
const DEFAULT_API_PORT = 9595;
const DEFAULT_WEB_UI_PORT = 80;
const DEFAULT_WEB_UI_HTTPS_PORT = 443;
let httpServer;
let webUiServer;
let currentWebUiPort = null; // Set when Web UI server starts
let webUiServerUsesHttps = false;
/** When `GSO_README_CAPTURE=1`, original prefs before forcing Web UI port 8765 for scripted screenshots */
let readmeCapturePrefsBackup = null;

// Cloudflare Quick Tunnel (WAN access) — child process + parsed trycloudflare.com URL
let cloudflaredProcess = null;
let tunnelUrl = null;
let cloudflaredKillTimer = null;
let tunnelQrWindow = null;
let tunnelQrHideTimer = null;

function startHttpServer() {
  httpServer = http.createServer(async (req, res) => {
    // Helpful request logging for diagnosing duplicate/looping calls
    try {
      const ua = (req.headers && req.headers['user-agent']) ? String(req.headers['user-agent']) : '';
      const from = (req.socket && req.socket.remoteAddress) ? String(req.socket.remoteAddress) : '';
      if (req.method !== 'OPTIONS') {
        // Very chatty: only emit in verbose mode
        logDebug(`[API] ${req.method} ${req.url} from ${from}${ua ? ` ua="${ua}"` : ''}`);
      }
    } catch (e) {
      // ignore logging failures
    }

    // Enable CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    
    if (req.method === 'OPTIONS') {
      res.writeHead(200);
      res.end();
      return;
    }

    // Controller allowlist: restrict who can call the API
    // If no controllerIps are configured, allow any client.
    try {
      const prefs = loadPreferences();
      if (!isControllerAllowedRequest(req, prefs)) {
        res.writeHead(403, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Forbidden' }));
        return;
      }
    } catch (e) {
      // If allowlist check fails unexpectedly, default to allowing (avoid breaking during startup)
    }

    // Path without query string (for routes that may be called with cache-busting query params)
    const apiReqPath = String(req.url || '').split('?')[0];
    
    // GET /api/status - Check if app is running and expose state for Companion variables/feedbacks
    if (req.method === 'GET' && apiReqPath === '/api/status') {
      (async () => {
        // Get login state and user info
        let loginState = false;
        let loggedInUser = null;
        try {
          const googleSession = getGoogleSession();
          const cookies = await googleSession.cookies.get({ domain: '.google.com' });
          const hasAuthCookies = cookies.some(cookie => 
            cookie.name === 'SID' || cookie.name === 'HSID' || cookie.name === 'SSID'
          );
          loginState = hasAuthCookies;
          
          if (hasAuthCookies) {
            // Try to get user email from cookies
            const emailCookie = cookies.find(cookie => 
              cookie.name === 'Email' || cookie.name === 'email' || 
              (cookie.value && cookie.value.includes('@'))
            );
            if (emailCookie && emailCookie.value && emailCookie.value.includes('@')) {
              loggedInUser = emailCookie.value;
            } else {
              // Try to get from any cookie value that looks like an email
              const emailLikeCookie = cookies.find(cookie => 
                cookie.value && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cookie.value)
              );
              if (emailLikeCookie) {
                loggedInUser = emailLikeCookie.value;
              }
            }
          }
        } catch (error) {
          console.error('[API] Error checking login state:', error);
        }
        
        const statusPrefs = loadPreferences();
        const state = {
          status: 'ok',
          version: appBuildInfo.version,
          buildNumber: appBuildInfo.buildNumber,
          presentationOpen: !!(presentationWindow && !presentationWindow.isDestroyed()),
          notesOpen: !!(notesWindow && !notesWindow.isDestroyed()),
          currentSlide: currentSlide,
          totalSlides: null,
          presentationUrl: lastPresentationUrl || null,
          slideInfo: null,
          isFirstSlide: null,
          isLastSlide: null,
          nextSlide: null,
          previousSlide: null,
          presentationTitle: null,
          timerElapsed: null,
          loginState: loginState,
          loggedInUser: loggedInUser || null,
          notesZoomSteps: notesZoomStepsFromDefault,
          notesZoomDefault: (statusPrefs.defaultNotesZoomSteps !== undefined && statusPrefs.defaultNotesZoomSteps !== null) ? clampNotesZoomSteps(statusPrefs.defaultNotesZoomSteps) : 0,
          notesLayout: statusPrefs.notesLayout || 'hide',
          runtime: {
            chrome: process.versions.chrome,
            electron: process.versions.electron,
            node: process.versions.node
          },
          presentationGpuMode: statusPrefs.presentationGpuMode || 'default',
          presentationNativeFullscreen: statusPrefs.presentationNativeFullscreen === true
        };
        
        // Get slide info and other data from notes window DOM
        if (notesWindow && !notesWindow.isDestroyed()) {
          try {
            const info = await notesWindow.webContents.executeJavaScript(`
              (function(){
                var result = {};
                
                // Get slide numbers from aria attributes
                var el = document.querySelector('[aria-posinset]');
                if (el) {
                  var cur = parseInt(el.getAttribute('aria-posinset'), 10);
                  var tot = parseInt(el.getAttribute('aria-setsize'), 10);
                  if (!isNaN(cur)) result.current = cur;
                  if (!isNaN(tot)) result.total = tot;
                }
                
                // Get presentation title from page title or DOM
                var titleEl = document.querySelector('title');
                if (titleEl) {
                  var titleText = titleEl.textContent;
                  // Extract title from "Presenter view - TITLE - Google Slides"
                  var match = titleText.match(/Presenter view - (.+?) - Google Slides/);
                  if (match) {
                    result.title = match[1];
                  } else {
                    result.title = titleText;
                  }
                }
                
                // Get timer value (look for timer display - usually shows "00:00:06" format)
                // Try to find elements containing time format
                var allText = document.body.innerText || document.body.textContent || '';
                var timeMatch = allText.match(/(\\d{1,2}:\\d{2}(?::\\d{2})?)/);
                if (timeMatch) {
                  result.timer = timeMatch[1];
                } else {
                  // Try specific timer elements
                  var timerEls = document.querySelectorAll('div, span');
                  for (var i = 0; i < timerEls.length; i++) {
                    var text = timerEls[i].textContent || timerEls[i].innerText || '';
                    var match = text.match(/^(\\d{1,2}:\\d{2}(?::\\d{2})?)$/);
                    if (match) {
                      result.timer = match[1];
                      break;
                    }
                  }
                }
                
                return result;
              })()
            `);
            
            if (info) {
              if (info.current != null) {
                state.currentSlide = info.current;
                // Calculate derived values
                if (info.total != null) {
                  state.totalSlides = info.total;
                  state.isFirstSlide = info.current === 1;
                  state.isLastSlide = info.current === info.total;
                  state.nextSlide = info.current < info.total ? info.current + 1 : null;
                  state.previousSlide = info.current > 1 ? info.current - 1 : null;
                  state.slideInfo = info.current + ' / ' + info.total;
                } else if (state.currentSlide !== null) {
                  // Use tracked currentSlide if DOM didn't provide total
                  state.isFirstSlide = state.currentSlide === 1;
                  state.nextSlide = state.currentSlide + 1;
                  state.previousSlide = state.currentSlide > 1 ? state.currentSlide - 1 : null;
                  if (state.totalSlides) {
                    state.isLastSlide = state.currentSlide === state.totalSlides;
                    state.slideInfo = state.currentSlide + ' / ' + state.totalSlides;
                  } else {
                    state.slideInfo = String(state.currentSlide);
                  }
                }
              }
              
              if (info.title) state.presentationTitle = info.title;
              if (info.timer) state.timerElapsed = info.timer;
            }
          } catch (e) { /* DOM not available or changed */ }
        }
        
        // Calculate derived values even if DOM didn't provide them
        if (state.currentSlide !== null && state.currentSlide !== undefined) {
          if (state.isFirstSlide === null) state.isFirstSlide = state.currentSlide === 1;
          if (state.nextSlide === null) state.nextSlide = state.currentSlide + 1;
          if (state.previousSlide === null) state.previousSlide = state.currentSlide > 1 ? state.currentSlide - 1 : null;
          if (state.slideInfo === null) {
            if (state.totalSlides) {
              state.slideInfo = state.currentSlide + ' / ' + state.totalSlides;
            } else {
              state.slideInfo = String(state.currentSlide);
            }
          }
          if (state.totalSlides && state.isLastSlide === null) {
            state.isLastSlide = state.currentSlide === state.totalSlides;
          }
        }
        
        // Get preferences for display IDs and backup controls
        const prefs = loadPreferences();
        state.presentationDisplayId = prefs.presentationDisplayId || null;
        state.notesDisplayId = prefs.notesDisplayId || null;
        state.notesEncodingIssue = lastNotesEncodingIssue;
        state.backupControlsEnabled = prefs.backupControlsEnabled !== false;
        state.tunnelEnabled = !!prefs.cloudflaredEnabled;
        state.tunnelUrl = tunnelUrl || null;
        state.tunnelQrVisible = !!(tunnelQrWindow && !tunnelQrWindow.isDestroyed());
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(state));
      })().catch(err => {
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: err.message }));
        }
      });
      return;
    }
    
    // POST /api/set-backup-controls - Enable or disable backup command forwarding (primary only)
    if (req.method === 'POST' && req.url === '/api/set-backup-controls') {
      let body = '';
      req.on('data', chunk => { body += chunk.toString(); });
      req.on('end', () => {
        try {
          const data = body ? JSON.parse(body) : {};
          const prefs = loadPreferences();
          if (prefs.primaryBackupMode !== 'primary') {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ backupControlsEnabled: prefs.backupControlsEnabled !== false, message: 'Not in primary mode' }));
            return;
          }
          const enabled = data.enabled !== false;
          prefs.backupControlsEnabled = enabled;
          savePreferences(prefs);
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ backupControlsEnabled: prefs.backupControlsEnabled, message: enabled ? 'Backup forwarding enabled' : 'Backup forwarding disabled' }));
        } catch (error) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message || 'Invalid request' }));
        }
      });
      return;
    }

    // POST /api/tunnel-enable - Start the Cloudflare Quick Tunnel
    if (req.method === 'POST' && req.url === '/api/tunnel-enable') {
      if (!isControllerAllowedRequest(req, loadPreferences())) {
        res.writeHead(403, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Forbidden' }));
        return;
      }
      const p = loadPreferences();
      p.cloudflaredEnabled = true;
      savePreferences(p);
      if (!cloudflaredProcess) startCloudflaredTunnel();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, message: 'Tunnel enabled' }));
      return;
    }

    // POST /api/tunnel-disable - Stop the Cloudflare Quick Tunnel
    if (req.method === 'POST' && req.url === '/api/tunnel-disable') {
      if (!isControllerAllowedRequest(req, loadPreferences())) {
        res.writeHead(403, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Forbidden' }));
        return;
      }
      const p = loadPreferences();
      p.cloudflaredEnabled = false;
      savePreferences(p);
      stopCloudflaredTunnel();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, message: 'Tunnel disabled' }));
      return;
    }

    // POST /api/show-tunnel-qr - Show QR overlay on notes display
    if (req.method === 'POST' && req.url === '/api/show-tunnel-qr') {
      if (!isControllerAllowedRequest(req, loadPreferences())) {
        res.writeHead(403, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Forbidden' }));
        return;
      }
      if (!tunnelUrl) {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: false, error: 'Tunnel not connected' }));
        return;
      }
      let body = '';
      req.on('data', chunk => { body += chunk.toString(); });
      req.on('end', () => {
        (async () => {
          try {
            const data = body ? JSON.parse(body) : {};
            const duration = Math.min(Math.max(parseInt(data.duration) || 20, 5), 300);
            await showTunnelQrOverlay(tunnelUrl, duration * 1000);
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'QR shown' }));
          } catch (error) {
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: error.message }));
          }
        })();
      });
      return;
    }

    // POST /api/hide-tunnel-qr - Dismiss QR overlay
    if (req.method === 'POST' && req.url === '/api/hide-tunnel-qr') {
      if (!isControllerAllowedRequest(req, loadPreferences())) {
        res.writeHead(403, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Forbidden' }));
        return;
      }
      hideTunnelQrOverlay();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, message: 'QR hidden' }));
      return;
    }

    // GET /api/backup-status - Get connection status of backup machines (primary mode only)
    if (req.method === 'GET' && req.url === '/api/backup-status') {
      (async () => {
        try {
          const status = await checkBackupStatus();
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(status));
        } catch (error) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        }
      })();
      return;
    }
    
    // GET /api/preferences - Get all preferences
    if (req.method === 'GET' && req.url === '/api/preferences') {
      try {
        const prefs = sanitizePreferencesForClient(loadPreferences());
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(prefs));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // POST /api/preferences - Save preferences
    if (req.method === 'POST' && req.url === '/api/preferences') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      
      req.on('end', () => {
        try {
          const data = JSON.parse(body);
          const prefs = loadPreferences();
          
          // Security: prevent changing controller allowlist via HTTP.
          // This setting is only editable from the desktop app UI (IPC).
          if (data && typeof data === 'object') {
            delete data.controllerIps;
            // Also desktop-only: web UI debug console gating
            delete data.webUiDebugConsoleEnabled;
            // WAN tunnel toggle is desktop-only (IPC); do not allow enabling via HTTP API
            delete data.cloudflaredEnabled;
            // Tunnel PIN is desktop-only (IPC)
            delete data.webUiTunnelPin;
            delete data.webUiTunnelPinClear;
            delete data.webUiTunnelPinScrypt;
            delete data.webUiTunnelPinSalt;
            delete data.webUiPinScope;
          }
          
          // Merge new preferences with existing ones
          Object.assign(prefs, data);
          const postSaveMeta = savePreferences(prefs) || {};

          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: true, message: 'Preferences saved', ...postSaveMeta }));
        } catch (error) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        }
      });
      return;
    }
    
    // GET /api/displays - Get available displays
    if (req.method === 'GET' && req.url === '/api/displays') {
      try {
        const displays = screen.getAllDisplays();
        const displayList = displays.map(display => ({
          id: display.id,
          bounds: display.bounds,
          label: `${display.bounds.width}x${display.bounds.height} @ (${display.bounds.x}, ${display.bounds.y})`,
          primary: display.id === screen.getPrimaryDisplay().id
        }));
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(displayList));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // POST /api/open-presentation - Open a presentation with URL
    if (req.method === 'POST' && req.url === '/api/open-presentation') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      
      req.on('end', async () => {
        try {
          const data = JSON.parse(body);
          const url = (data.url || '').trim();
          
          if (!url) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'URL is required' }));
            return;
          }
          
          console.log('[API] Opening presentation:', url);
          
          // Close any existing presentation windows
          try {
            if (notesWindow && !notesWindow.isDestroyed()) {
              console.log('[API] Closing existing notes window');
              notesWindow.removeAllListeners('closed');
              notesWindow.close();
              notesWindow = null;
            }
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Closing existing presentation window');
              presentationWindow.removeAllListeners('closed');
              presentationWindow.close();
              presentationWindow = null;
            }
            currentSlide = null;
          } catch (error) {
            console.error('[API] Error closing existing windows:', error.message);
          }
          
          // Load preferences for monitor selection
          const prefs = loadPreferences();
          const displays = screen.getAllDisplays();
          
          const presentationDisplayId = Number(prefs.presentationDisplayId);
          const notesDisplayId = Number(prefs.notesDisplayId);
          
          const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
          const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];
          
          console.log('[API] Using presentation display:', presentationDisplay.id);
          console.log('[API] Using notes display:', notesDisplay.id);
          
          // Open the presentation using the same logic as the IPC handler
          // Create the presentation window
          // Note: Don't use fullscreen: true in constructor as it creates a new Space on macOS
          // We'll use setSimpleFullScreen() after creation to avoid Spaces conflicts
          presentationWindow = new BrowserWindow(getPresentationBrowserWindowOptions(presentationDisplay.bounds));
          
          // Set simple fullscreen on macOS to avoid Spaces conflicts
          if (process.platform === 'darwin') {
            applyPresentationFullscreenChrome(presentationWindow, prefs);
          }
          attachCrashHandlers(presentationWindow, 'presentation');
          
          // Set up window open handler for speaker notes popup (open at notes-display size for narrow-preview layout)
          presentationWindow.webContents.setWindowOpenHandler(({ url }) => {
            const windowOptions = getSpeakerNotesWindowOptions(notesDisplay);
            return { action: 'allow', overrideBrowserWindowOptions: windowOptions };
          });
          
          // Listen for notes window creation
          const windowCreatedListener = (event, window) => {
            if (window !== presentationWindow && window !== mainWindow) {
              console.log('[API] Notes window created');
              notesWindow = window;
              onNotesWindowCreated(window);
              attachCrashHandlers(window, 'notes');
              // Add Escape key handler to notes window
              window.webContents.on('before-input-event', (event, input) => {
                if (input.key === 'Escape' && input.type === 'keyDown') {
                  console.log('[API] Escape pressed in notes window, closing all windows');
                  event.preventDefault();
                  if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
                  if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
                }
              });

              applySpeakerNotesInitialGeometry(window, notesDisplay, null);

              app.removeListener('browser-window-created', windowCreatedListener);
            }
          };
          app.on('browser-window-created', windowCreatedListener);

          // Navigation listener (no auto-launch of notes - user must manually start notes)
          const navigationListener = async (event, navUrl) => {
            console.log('[API] Navigated to:', navUrl);
            // Just log navigation, don't auto-launch notes
          };
          
          presentationWindow.webContents.on('did-navigate', navigationListener);
          
          // Listen for page load
          presentationWindow.webContents.once('did-finish-load', async () => {
            console.log('[API] Page finished loading');
            if (!presentationWindow || presentationWindow.isDestroyed()) {
              console.log('[API] Window destroyed before processing');
              return;
            }

            // We now load /present directly (via toPresentUrl), so we should NOT press Ctrl+Shift+F5 here.
            // Pressing it can cause flaky behavior (extra reloads / exits) depending on Slides state.
          });
          
          presentationWindow.on('closed', () => {
            presentationWindow = null;
            currentSlide = null;
          });
          
          // Escape key handler for presentation window
          presentationWindow.webContents.on('before-input-event', (event, input) => {
          if (input.key === 'Escape' && input.type === 'keyDown') {
            event.preventDefault();
            if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
            if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
          }
          });
          
          const presentUrl = toPresentUrl(url);
          console.log('[API] Loading PRESENT URL:', presentUrl);
          lastPresentationUrl = url; // Store original URL (not /present URL) for reload
          currentSlide = 1;
          resetNotesZoomForNewPresentation();
          presentationWindow.loadURL(presentUrl);
          presentationWindow.show();
          
          // Broadcast to backups (async, don't wait)
          sendToBackups('/api/open-presentation', { url: url }).catch(err => {
            console.error('[Backup] Error broadcasting open-presentation:', err);
          });
          
          // Send response immediately
          if (!res.headersSent) {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Presentation opened (notes not auto-started)' }));
          }
        } catch (error) {
          console.error('[API] Error:', error);
          if (!res.headersSent) {
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: error.message }));
          }
        }
      });
      return;
    }

    // POST /api/open-presentation-with-notes - Open a presentation and automatically launch speaker notes
    if (req.method === 'POST' && req.url === '/api/open-presentation-with-notes') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      
      req.on('end', async () => {
        try {
          const data = JSON.parse(body);
          const { url } = data;
          
          if (!url) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'URL is required' }));
            return;
          }
          
          console.log('[API] Opening presentation with notes:', url);
          
          // Close any existing presentation windows
          try {
            if (notesWindow && !notesWindow.isDestroyed()) {
              console.log('[API] Closing existing notes window');
              notesWindow.removeAllListeners('closed');
              notesWindow.close();
              notesWindow = null;
            }
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Closing existing presentation window');
              presentationWindow.removeAllListeners('closed');
              presentationWindow.close();
              presentationWindow = null;
            }
            currentSlide = null;
          } catch (error) {
            console.error('[API] Error closing existing windows:', error.message);
          }
          
          // Load preferences for monitor selection
          const prefs = loadPreferences();
          const displays = screen.getAllDisplays();
          
          const presentationDisplayId = Number(prefs.presentationDisplayId);
          const notesDisplayId = Number(prefs.notesDisplayId);
          
          const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
          const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];
          
          console.log('[API] Using presentation display:', presentationDisplay.id);
          console.log('[API] Using notes display:', notesDisplay.id);
          
          // Create the presentation window
          // Note: Don't use fullscreen: true in constructor as it creates a new Space on macOS
          // We'll use setSimpleFullScreen() after creation to avoid Spaces conflicts
          presentationWindow = new BrowserWindow(getPresentationBrowserWindowOptions(presentationDisplay.bounds));
          
          // Set simple fullscreen on macOS to avoid Spaces conflicts
          if (process.platform === 'darwin') {
            applyPresentationFullscreenChrome(presentationWindow, prefs);
          }
          attachCrashHandlers(presentationWindow, 'presentation');
          
          // Set up window open handler for speaker notes popup (open at notes-display size for narrow-preview layout)
          presentationWindow.webContents.setWindowOpenHandler(({ url }) => {
            const windowOptions = getSpeakerNotesWindowOptions(notesDisplay);
            return { action: 'allow', overrideBrowserWindowOptions: windowOptions };
          });
          
          // Listen for notes window creation
          const windowCreatedListener = (event, window) => {
            if (window !== presentationWindow && window !== mainWindow) {
              console.log('[API] Notes window created');
              notesWindow = window;
              onNotesWindowCreated(window);
              attachCrashHandlers(window, 'notes');
              window.webContents.on('before-input-event', (event, input) => {
                if (input.key === 'Escape' && input.type === 'keyDown') {
                  event.preventDefault();
                  if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
                  if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
                }
              });
              applySpeakerNotesInitialGeometry(window, notesDisplay, null);

              app.removeListener('browser-window-created', windowCreatedListener);
            }
          };
          app.on('browser-window-created', windowCreatedListener);

          // Auto-launch speaker notes reliably (some decks load fast and can miss a single 's' press).
          // We'll retry a few times until the notes window is created.
          let notesAttempts = 0;
          const maxNotesAttempts = 8;
          let notesRetryTimer = null;

          const sendSpeakerNotesKey = async (reason) => {
            if (!presentationWindow || presentationWindow.isDestroyed()) return false;
            if (notesWindow && !notesWindow.isDestroyed()) return true;
            if (notesAttempts >= maxNotesAttempts) return false;

            notesAttempts += 1;
            try {
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 80));
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
              presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
              console.log(`[API] Speaker notes attempt ${notesAttempts}/${maxNotesAttempts} (${reason}) - sent "s" key`);
            } catch (error) {
              console.error('[API] Error sending "s" key for speaker notes:', error);
            }

            // Schedule another attempt if notes window hasn't appeared yet
            if (!notesWindow || notesWindow.isDestroyed()) {
              if (notesRetryTimer) clearTimeout(notesRetryTimer);
              notesRetryTimer = setTimeout(() => {
                sendSpeakerNotesKey('retry');
              }, 700);
            }

            return true;
          };

          const navigationListener = async (event, navUrl) => {
            // Check if we're in presentation mode (URL contains /present/ or /localpresent but not /presentation/)
            const isPresentMode = (navUrl.includes('/present/') || navUrl.includes('/localpresent')) && !navUrl.includes('/presentation/');
            if (isPresentMode) {
              // Slight delay to allow the presentation UI to become interactive
              await new Promise(resolve => setTimeout(resolve, 250));
              await sendSpeakerNotesKey('did-navigate');
            }
          };

          presentationWindow.webContents.on('did-navigate', navigationListener);
          
          // Listen for page load
          presentationWindow.webContents.once('did-finish-load', async () => {
            console.log('[API] Page finished loading');
            if (!presentationWindow || presentationWindow.isDestroyed()) {
              console.log('[API] Window destroyed before processing');
              return;
            }

            // If we're NOT already in /present or /localpresent, attempt to trigger present mode.
            // (Most of the time we load /present directly, so this will be skipped.)
            try {
              const loadedUrl = presentationWindow.webContents.getURL() || '';
              const isPresentAlready = loadedUrl.includes('/present/') || loadedUrl.includes('/localpresent');
              if (!isPresentAlready) {
                await new Promise(resolve => setTimeout(resolve, 200));
                if (presentationWindow && !presentationWindow.isDestroyed()) {
                  console.log('[API] Not in present mode yet, triggering Ctrl+Shift+F5...');
                  presentationWindow.focus();
                  await new Promise(resolve => setTimeout(resolve, 80));
                  presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
                  presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
                }
              }
            } catch (e) {
              // ignore
            }

            // Always attempt notes after load (covers cases where did-navigate isn't fired as expected).
            setTimeout(() => {
              sendSpeakerNotesKey('did-finish-load');
            }, 650);
          });
          
          presentationWindow.on('closed', () => {
            presentationWindow = null;
            currentSlide = null;
          });
          
          // Escape key handler for presentation window
          presentationWindow.webContents.on('before-input-event', (event, input) => {
          if (input.key === 'Escape' && input.type === 'keyDown') {
            event.preventDefault();
            if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
            if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
          }
          });
          
          const presentUrl = toPresentUrl(url);
          console.log('[API] Loading PRESENT URL:', presentUrl);
          lastPresentationUrl = url; // Store original URL (not /present URL) for reload
          currentSlide = 1;
          resetNotesZoomForNewPresentation();
          presentationWindow.loadURL(presentUrl);
          presentationWindow.show();
          
          // Broadcast to backups (async, don't wait)
          sendToBackups('/api/open-presentation-with-notes', { url: url }).catch(err => {
            console.error('[Backup] Error broadcasting open-presentation-with-notes:', err);
          });
          
          // Send response immediately
          if (!res.headersSent) {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Presentation opened with notes' }));
          }
        } catch (error) {
          console.error('[API] Error opening presentation with notes:', error);
          if (!res.headersSent) {
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: error.message }));
          }
        }
      });
      return;
    }
    
    // POST /api/close-presentation - Close current presentation
    if (req.method === 'POST' && req.url === '/api/close-presentation') {
      console.log('[API] Closing presentation');
      
      // Broadcast to backups (async, don't wait)
      sendToBackups('/api/close-presentation', {}).catch(err => {
        console.error('[Backup] Error broadcasting close-presentation:', err);
      });
      
      // Send response first before closing windows
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, message: 'Presentation closed' }));
      
      // Close windows after sending response to avoid errors
      setImmediate(() => {
        try {
          if (notesWindow && !notesWindow.isDestroyed()) {
            notesWindow.removeAllListeners('closed');
            notesWindow.close();
            notesWindow = null;
          }
          if (presentationWindow && !presentationWindow.isDestroyed()) {
            presentationWindow.removeAllListeners('closed');
            presentationWindow.close();
            presentationWindow = null;
          }
          currentSlide = null;
        } catch (error) {
          console.error('[API] Error closing windows:', error.message);
        }
      });
      
      return;
    }
    
    // POST /api/next-slide - Go to next slide
    if (req.method === 'POST' && req.url === '/api/next-slide') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }
      
      try {
        presentationWindow.focus();
        presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Right' });
        presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Right' });
        currentSlide = (typeof currentSlide === 'number' ? currentSlide + 1 : 1);
        
        // Broadcast to backups (async, don't wait)
        sendToBackups('/api/next-slide', {}).catch(err => {
          console.error('[Backup] Error broadcasting next-slide:', err);
        });
        
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Next slide' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // POST /api/previous-slide - Go to previous slide
    if (req.method === 'POST' && req.url === '/api/previous-slide') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }
      
      try {
        presentationWindow.focus();
        presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Left' });
        presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Left' });
        currentSlide = (typeof currentSlide === 'number' && currentSlide > 1 ? currentSlide - 1 : 1);
        
        // Broadcast to backups (async, don't wait)
        sendToBackups('/api/previous-slide', {}).catch(err => {
          console.error('[Backup] Error broadcasting previous-slide:', err);
        });
        
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Previous slide' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/go-to-slide - Navigate to a specific slide number
    if (req.method === 'POST' && req.url === '/api/go-to-slide') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }

      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });

      req.on('end', async () => {
        try {
          const data = JSON.parse(body);
          const targetSlide = parseInt(data.slide, 10);

          if (isNaN(targetSlide) || targetSlide < 1) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'Valid slide number (>= 1) is required' }));
            return;
          }

          // Get current slide (from our tracking or default to 1)
          const current = typeof currentSlide === 'number' ? currentSlide : 1;
          const slidesToMove = targetSlide - current;

          if (slidesToMove === 0) {
            // Broadcast to backups even if already on target slide (for sync)
            sendToBackups('/api/go-to-slide', { slide: targetSlide }).catch(err => {
              console.error('[Backup] Error broadcasting go-to-slide:', err);
            });
            
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Already on slide ' + targetSlide }));
            return;
          }

          presentationWindow.focus();
          await new Promise(resolve => setTimeout(resolve, 50));

          // Send arrow key presses to navigate
          const keyCode = slidesToMove > 0 ? 'Right' : 'Left';
          const count = Math.abs(slidesToMove);

          for (let i = 0; i < count; i++) {
            presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: keyCode });
            presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: keyCode });
            // Small delay between key presses to ensure they're processed
            if (i < count - 1) {
              await new Promise(resolve => setTimeout(resolve, 100));
            }
          }

          // Update our tracking
          currentSlide = targetSlide;
          
          // Broadcast to backups (async, don't wait)
          sendToBackups('/api/go-to-slide', { slide: targetSlide }).catch(err => {
            console.error('[Backup] Error broadcasting go-to-slide:', err);
          });

          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ 
            success: true, 
            message: `Navigated to slide ${targetSlide}`,
            fromSlide: current,
            toSlide: targetSlide
          }));
        } catch (error) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        }
      });
      return;
    }

    // POST /api/reload-presentation - Close, reopen, and return to current slide
    if (req.method === 'POST' && req.url === '/api/reload-presentation') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }

      if (!lastPresentationUrl) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No previous presentation URL stored' }));
        return;
      }

      // Send response immediately (this will be async)
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, message: 'Reloading presentation...' }));

      // Do the reload asynchronously (reuses shared reopenPresentationAtSlide)
      (async () => {
        try {
          const savedSlide = typeof currentSlide === 'number' ? currentSlide : 1;
          const notesWereOpen = !!(notesWindow && !notesWindow.isDestroyed());
          let savedNotesBounds = null;
          if (notesWereOpen && notesWindow && !notesWindow.isDestroyed()) {
            try {
              savedNotesBounds = notesWindow.getBounds();
              console.log('[API] Reload: Cached notes window bounds', savedNotesBounds);
            } catch (e) {
              console.warn('[API] Reload: Could not cache notes bounds:', e.message);
            }
            await logNotesZoomFontProbeBeforeReload(notesWindow, 'reload');
          }
          const savedNotesZoomSteps = notesWereOpen ? notesZoomStepsFromDefault : getDefaultNotesZoomStepsFromPrefs();
          const urlToReload = lastPresentationUrl;
          console.log('[API] Reload: Saving state - slide:', savedSlide, 'notes open:', notesWereOpen, 'notesZoomSteps:', savedNotesZoomSteps, 'URL:', urlToReload);
          await reopenPresentationAtSlide(urlToReload, savedSlide, notesWereOpen, savedNotesBounds, savedNotesZoomSteps);
          console.log('[API] Reload: Complete');
        } catch (error) {
          console.error('[API] Error during reload:', error);
        }
      })();

      return;
    }
    
    // POST /api/toggle-video - Toggle video playback (k key)
    if (req.method === 'POST' && req.url === '/api/toggle-video') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }
      
      try {
        presentationWindow.focus();
        presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'K' });
        presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 'k' });
        presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'K' });
        
        // Broadcast to backups (async, don't wait)
        sendToBackups('/api/toggle-video', {}).catch(err => {
          console.error('[Backup] Error broadcasting toggle-video:', err);
        });
        
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Video toggled' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/open-speaker-notes - Open/start speaker notes (s key)
    if (req.method === 'POST' && req.url === '/api/open-speaker-notes') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }

      try {
        console.log('[API] Opening speaker notes');
        presentationWindow.focus();
        await new Promise(resolve => setTimeout(resolve, 50));
        presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
        presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
        presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
        
        // Broadcast to backups (async, don't wait)
        sendToBackups('/api/open-speaker-notes', {}).catch(err => {
          console.error('[Backup] Error broadcasting open-speaker-notes:', err);
        });

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Speaker notes opened' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/close-speaker-notes - Close the speaker notes window
    if (req.method === 'POST' && req.url === '/api/close-speaker-notes') {
      if (!notesWindow || notesWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }

      try {
        notesWindow.close();
        notesWindow = null;
        
        // Broadcast to backups (async, don't wait)
        sendToBackups('/api/close-speaker-notes', {}).catch(err => {
          console.error('[Backup] Error broadcasting close-speaker-notes:', err);
        });
        
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Speaker notes closed' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/relaunch-speaker-notes - Close and reopen notes window to apply layout changes
    if (req.method === 'POST' && req.url === '/api/relaunch-speaker-notes') {
      try {
        const result = await relaunchSpeakerNotesWindow();
        if (!result.success) {
          res.writeHead(404, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: result.error || 'Relaunch failed' }));
          return;
        }
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: result.message || 'Speaker notes relaunched' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/scroll-notes-down - Scroll speaker notes down (JS only, no keyboard)
    if (req.method === 'POST' && req.url === '/api/scroll-notes-down') {
      if (!notesWindow || notesWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }
      try {
        notesWindow.webContents.executeJavaScript(`
          (function() {
            // Find scrollable elements - try common patterns in Google Slides presenter view
            var scrollable = null;
            var allElements = document.querySelectorAll('*');
            for (var i = 0; i < allElements.length; i++) {
              var el = allElements[i];
              var style = window.getComputedStyle(el);
              if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && 
                  el.scrollHeight > el.clientHeight) {
                scrollable = el;
                break;
              }
            }
            // Fallback: try document body or documentElement if they're scrollable
            if (!scrollable) {
              if (document.body && document.body.scrollHeight > document.body.clientHeight) {
                scrollable = document.body;
              } else if (document.documentElement && document.documentElement.scrollHeight > document.documentElement.clientHeight) {
                scrollable = document.documentElement;
              }
            }
            if (scrollable) {
              scrollable.scrollBy(0, 150);
              return { success: true, scrolled: true };
            }
            return { success: false, error: 'No scrollable element found' };
          })()
        `).then(result => {
          if (result.success && result.scrolled) {
            // Broadcast to backups (async, don't wait)
            sendToBackups('/api/scroll-notes-down', {}).catch(err => {
              console.error('[Backup] Error broadcasting scroll-notes-down:', err);
            });
            
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Notes scrolled down' }));
          } else {
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: result.error || 'Could not scroll notes' }));
          }
        }).catch(error => {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        });
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/scroll-notes-up - Scroll speaker notes up (JS only, no keyboard)
    if (req.method === 'POST' && req.url === '/api/scroll-notes-up') {
      if (!notesWindow || notesWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }
      try {
        notesWindow.webContents.executeJavaScript(`
          (function() {
            // Find scrollable elements - try common patterns in Google Slides presenter view
            var scrollable = null;
            var allElements = document.querySelectorAll('*');
            for (var i = 0; i < allElements.length; i++) {
              var el = allElements[i];
              var style = window.getComputedStyle(el);
              if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && 
                  el.scrollHeight > el.clientHeight) {
                scrollable = el;
                break;
              }
            }
            // Fallback: try document body or documentElement if they're scrollable
            if (!scrollable) {
              if (document.body && document.body.scrollHeight > document.body.clientHeight) {
                scrollable = document.body;
              } else if (document.documentElement && document.documentElement.scrollHeight > document.documentElement.clientHeight) {
                scrollable = document.documentElement;
              }
            }
            if (scrollable) {
              scrollable.scrollBy(0, -150);
              return { success: true, scrolled: true };
            }
            return { success: false, error: 'No scrollable element found' };
          })()
        `).then(result => {
          if (result.success && result.scrolled) {
            // Broadcast to backups (async, don't wait)
            sendToBackups('/api/scroll-notes-up', {}).catch(err => {
              console.error('[Backup] Error broadcasting scroll-notes-up:', err);
            });
            
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Notes scrolled up' }));
          } else {
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: result.error || 'Could not scroll notes' }));
          }
        }).catch(error => {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        });
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // POST /api/zoom-in-notes - Zoom in on speaker notes
    if (req.method === 'POST' && req.url === '/api/zoom-in-notes') {
      console.log('[API] Zoom in on speaker notes requested');
      
      if (!notesWindow || notesWindow.isDestroyed()) {
        console.log('[API] No speaker notes window is open');
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }
      
      try {
        executeNotesZoomClickIn(notesWindow).then(result => {
          if (result.success) {
            notesZoomStepsFromDefault = clampNotesZoomSteps(notesZoomStepsFromDefault + 1);
            console.log('[API] ✓ Dispatched mouse events to zoom in button');
            
            // Broadcast to backups (async, don't wait)
            sendToBackups('/api/zoom-in-notes', {}).catch(err => {
              console.error('[Backup] Error broadcasting zoom-in-notes:', err);
            });
            
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Zoomed in on notes' }));
          } else {
            console.log('[API] ✗ Zoom in button not found');
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: result.error }));
          }
        }).catch(error => {
          console.error('[API] Error executing zoom in script:', error.message);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        });
      } catch (error) {
        console.error('[API] Error zooming in on notes:', error.message);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // POST /api/zoom-out-notes - Zoom out on speaker notes
    if (req.method === 'POST' && req.url === '/api/zoom-out-notes') {
      console.log('[API] Zoom out on speaker notes requested');
      
      if (!notesWindow || notesWindow.isDestroyed()) {
        console.log('[API] No speaker notes window is open');
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }
      
      try {
        executeNotesZoomClickOut(notesWindow).then(result => {
          if (result.success) {
            notesZoomStepsFromDefault = clampNotesZoomSteps(notesZoomStepsFromDefault - 1);
            console.log('[API] ✓ Dispatched mouse events to zoom out button');
            
            // Broadcast to backups (async, don't wait)
            sendToBackups('/api/zoom-out-notes', {}).catch(err => {
              console.error('[Backup] Error broadcasting zoom-out-notes:', err);
            });
            
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Zoomed out on notes' }));
          } else {
            console.log('[API] ✗ Zoom out button not found');
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: result.error }));
          }
        }).catch(error => {
          console.error('[API] Error executing zoom out script:', error.message);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        });
      } catch (error) {
        console.error('[API] Error zooming out on notes:', error.message);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // GET /api/get-speaker-notes - Get current speaker notes content
    if (req.method === 'GET' && apiReqPath === '/api/get-speaker-notes') {
      if (!notesWindow || notesWindow.isDestroyed()) {
        res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
        res.end(JSON.stringify({ success: false, notes: '', error: 'No speaker notes window is open' }));
        return;
      }

      (async () => {
        try {
          const notesContent = await notesWindow.webContents.executeJavaScript(`
            (function(){
              // Only the scrollable notes body - no tabs, no "AUDIENCE TOOLS" / "Speaker Notes"
              var el = document.querySelector('div.punch-viewer-speakernotes-text-body-scrollable');
              if (!el) return 'No notes available for this slide.';
              var raw = (el.innerText || el.textContent || '').trim();
              return raw.length > 0 ? raw : 'No notes available for this slide.';
            })()
          `);
          const rawNotes = notesContent || 'No notes available for this slide.';
          logFirstReplacementCharContext(rawNotes, 'raw ');
          const replacementCharsFound = countReplacementChars(rawNotes);
          const encodingIssuesDetected = replacementCharsFound > 0;
          lastNotesEncodingIssue = encodingIssuesDetected;
          const notes = normalizeSpeakerNotes(rawNotes) || 'No notes available for this slide.';
          res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ success: true, notes, encodingIssuesDetected, replacementCharsFound }), 'utf8');
        } catch (error) {
          console.error('[API] Error getting speaker notes:', error);
          res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Access-Control-Allow-Origin': '*' });
          res.end(JSON.stringify({ success: false, notes: '', error: error.message }), 'utf8');
        }
      })();
      return;
    }

    // GET /api/get-slide-previews - Get current + next slide preview images (from presenter view)
    if (req.method === 'GET' && req.url === '/api/get-slide-previews') {
      if (!notesWindow || notesWindow.isDestroyed()) {
        res.writeHead(200, {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        });
        res.end(JSON.stringify({ success: false, error: 'No speaker notes window is open' }));
        return;
      }

      (async () => {
        try {
          const result = await captureSlidePreviewsFromNotesWindow({ maxSize: 200 });
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(result));
        } catch (error) {
          console.error('[API] Error getting slide previews:', error);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, error: error.message }));
        }
      })();
      return;
    }
    
    // GET /api/get-stagetimer-status - Get live timer data from stagetimer.io
    if (req.method === 'GET' && req.url === '/api/get-stagetimer-status') {
      const prefs = loadPreferences();
      const roomId = prefs.stagetimerRoomId;
      const apiKey = prefs.stagetimerApiKey;
      
      if (!roomId || !apiKey) {
        res.writeHead(200, { 
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        });
        res.end(JSON.stringify({ 
          success: false, 
          error: 'Stagetimer not configured. Please set Room ID and API Key in Settings.',
          configured: false
        }));
        return;
      }
      
      // Call stagetimer.io API for status
      const statusUrl = `https://api.stagetimer.io/v1/get_status?room_id=${encodeURIComponent(roomId)}&api_key=${encodeURIComponent(apiKey)}`;
      const messagesUrl = `https://api.stagetimer.io/v1/get_all_messages?room_id=${encodeURIComponent(roomId)}&api_key=${encodeURIComponent(apiKey)}`;
      
      // Fetch status and messages in parallel, then fetch timer using timer_id
      let statusData = null;
      let messagesData = null;
      let timerData = null;
      let completed = 0;
      let timerCompleted = false;
      let timerTimeout = null;
      const totalRequests = 2; // Status and messages first
      const TIMER_FETCH_TIMEOUT = 5000; // 5 second timeout for timer fetch
      
      function sendResponse() {
        // Wait for status and messages
        if (completed < totalRequests) return;
        // If we're waiting for timer, give it a chance, but don't wait forever
        if (timerData === null && !timerCompleted) {
          // Set a timeout to send response even if timer fetch hangs
          if (!timerTimeout) {
            timerTimeout = setTimeout(() => {
              console.warn('[API] Timer fetch timeout, sending response without timer data');
              timerCompleted = true;
              timerData = { ok: false, data: {}, error: 'Timeout' };
              sendResponse();
            }, TIMER_FETCH_TIMEOUT);
          }
          return;
        }
        
        // Clear timeout if we got here normally
        if (timerTimeout) {
          clearTimeout(timerTimeout);
          timerTimeout = null;
        }
        
        try {
          if (statusData && statusData.ok && statusData.data) {
            const status = statusData.data;
            const now = status.server_time || Date.now();
            
            // Calculate remaining/elapsed time
            let remainingMs = 0;
            let elapsedMs = 0;
            let displayTime = '0:00';
            let isRunning = status.running || false;
            
            if (status.finish && status.start) {
              const duration = status.finish - status.start;
              
              if (isRunning) {
                remainingMs = status.finish - now; // Allow negative values
                elapsedMs = now - status.start;
              } else if (status.pause) {
                // Timer is paused
                elapsedMs = status.pause - status.start;
                remainingMs = duration - elapsedMs;
              } else {
                // Timer not started
                remainingMs = duration;
                elapsedMs = 0;
              }
              
              // Format time as MM:SS or HH:MM:SS (allow negative)
              const totalSeconds = Math.floor(remainingMs / 1000);
              const isNegative = totalSeconds < 0;
              const absSeconds = Math.abs(totalSeconds);
              const hours = Math.floor(absSeconds / 3600);
              const minutes = Math.floor((absSeconds % 3600) / 60);
              const seconds = absSeconds % 60;
              
              const sign = isNegative ? '-' : '';
              const minStr = String(minutes).padStart(2, '0');
              const secStr = String(seconds).padStart(2, '0');
              
              if (hours > 0) {
                displayTime = sign + hours + ':' + minStr + ':' + secStr;
              } else {
                displayTime = sign + minutes + ':' + secStr;
              }
            }
            
            // Process messages
            let activeMessages = [];
            if (messagesData && messagesData.ok && messagesData.data) {
              // Check if messages is an array directly or nested in data.messages
              let messages = [];
              if (Array.isArray(messagesData.data)) {
                messages = messagesData.data;
              } else if (messagesData.data.messages && Array.isArray(messagesData.data.messages)) {
                messages = messagesData.data.messages;
              }
              
              console.log('[API] Processing messages, found:', messages.length, 'total messages');
              activeMessages = messages
                .filter(msg => msg && msg.showing === true)
                .map(msg => ({
                  text: msg.text || '',
                  color: msg.color || 'white',
                  bold: msg.bold || false,
                  uppercase: msg.uppercase || false
                }));
              console.log('[API] Active messages:', activeMessages.length);
            } else {
              console.log('[API] No messages data or not ok:', messagesData);
            }
            
            // Get timer name and speaker from timer data
            let timerName = '';
            let speakerName = '';
            console.log('[API] Processing timer data, timerData:', timerData ? 'exists' : 'null');
            console.log('[API] timerData.ok:', timerData?.ok);
            console.log('[API] timerData.data:', timerData?.data ? JSON.stringify(timerData.data) : 'null');
            if (timerData && timerData.ok && timerData.data) {
              timerName = timerData.data.name || '';
              speakerName = timerData.data.speaker || '';
              console.log('[API] Extracted timerName:', timerName, 'speakerName:', speakerName);
            } else {
              console.log('[API] Timer data not available or invalid. timerData:', timerData);
            }
            
            res.writeHead(200, { 
              'Content-Type': 'application/json',
              'Access-Control-Allow-Origin': '*'
            });
            res.end(JSON.stringify({
              success: true,
              configured: true,
              running: isRunning,
              displayTime: displayTime,
              remainingMs: remainingMs,
              elapsedMs: elapsedMs,
              timerId: status.timer_id || null,
              start: status.start,
              finish: status.finish,
              pause: status.pause,
              serverTime: status.server_time,
              messages: activeMessages,
              timerName: timerName,
              speaker: speakerName,
              // Debug info
              _debug: {
                timerDataOk: timerData?.ok || false,
                timerDataExists: !!timerData,
                timerDataMessage: timerData?.message || null,
                rawTimerName: timerData?.data?.name || null,
                rawSpeaker: timerData?.data?.speaker || null
              }
            }));
          } else {
            // Status fetch failed or returned invalid data
            // Still try to send response with whatever we have
            res.writeHead(200, { 
              'Content-Type': 'application/json',
              'Access-Control-Allow-Origin': '*'
            });
            const errorMsg = statusData 
              ? (statusData.message || statusData.error || 'Failed to get timer status')
              : 'Failed to get timer status';
            console.error('[API] Status fetch failed:', errorMsg, 'statusData:', statusData);
            res.end(JSON.stringify({ 
              success: false, 
              error: errorMsg,
              configured: true,
              _debug: {
                statusDataExists: !!statusData,
                statusDataOk: statusData?.ok,
                statusDataMessage: statusData?.message,
                timerDataExists: !!timerData,
                timerDataOk: timerData?.ok
              }
            }));
          }
        } catch (error) {
          console.error('[API] Error processing stagetimer response:', error);
          res.writeHead(200, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ 
            success: false, 
            error: 'Failed to process response: ' + error.message,
            configured: true
          }));
        }
      }
      
      // Fetch status
      const statusReq = https.get(statusUrl, (apiRes) => {
        let data = '';
        
        // Check HTTP status code
        if (apiRes.statusCode !== 200) {
          console.error('[API] Stagetimer status HTTP error:', apiRes.statusCode, apiRes.statusMessage);
          apiRes.on('data', () => {}); // Drain response
          apiRes.on('end', () => {
            statusData = { ok: false, message: `HTTP ${apiRes.statusCode}: ${apiRes.statusMessage}` };
            completed++;
            sendResponse();
          });
          return;
        }
        
        apiRes.on('data', (chunk) => {
          data += chunk;
        });
        
        apiRes.on('end', () => {
          try {
            statusData = JSON.parse(data);
            if (!statusData || !statusData.ok) {
              console.error('[API] Stagetimer status API error:', statusData?.message || 'Unknown error');
            }
            completed++;
            
            // After getting status, fetch timer using timer_id if available
            if (statusData && statusData.ok && statusData.data && statusData.data.timer_id) {
              const timerId = statusData.data.timer_id;
              const timerUrl = `https://api.stagetimer.io/v1/get_timer?room_id=${encodeURIComponent(roomId)}&api_key=${encodeURIComponent(apiKey)}&timer_id=${encodeURIComponent(timerId)}`;
              
              console.log('[API] Fetching timer with timer_id:', timerId);
              const timerReq = https.get(timerUrl, (timerRes) => {
                let timerDataStr = '';
                
                if (timerRes.statusCode !== 200) {
                  console.error('[API] Stagetimer timer HTTP error:', timerRes.statusCode, timerRes.statusMessage);
                  timerRes.on('data', () => {}); // Drain response
                  timerRes.on('end', () => {
                    if (!timerCompleted) {
                      timerData = { ok: false, message: `HTTP ${timerRes.statusCode}: ${timerRes.statusMessage}`, data: {} };
                      timerCompleted = true;
                      if (timerTimeout) clearTimeout(timerTimeout);
                      sendResponse();
                    }
                  });
                  return;
                }
                
                timerRes.on('data', (chunk) => {
                  timerDataStr += chunk;
                });
                
                timerRes.on('end', () => {
                  if (timerCompleted) return; // Already handled by timeout
                  try {
                    timerData = JSON.parse(timerDataStr);
                    console.log('[API] Stagetimer timer response:', JSON.stringify(timerData, null, 2));
                    console.log('[API] Timer data.name:', timerData?.data?.name);
                    console.log('[API] Timer data.speaker:', timerData?.data?.speaker);
                    timerCompleted = true;
                    if (timerTimeout) clearTimeout(timerTimeout);
                    sendResponse();
                  } catch (error) {
                    console.error('[API] Error parsing stagetimer timer response:', error);
                    console.error('[API] Raw timer response:', timerDataStr);
                    timerData = { ok: false, data: {}, error: error.message };
                    timerCompleted = true;
                    if (timerTimeout) clearTimeout(timerTimeout);
                    sendResponse();
                  }
                });
              });
              
              timerReq.on('error', (error) => {
                if (timerCompleted) return; // Already handled by timeout
                console.error('[API] Error calling stagetimer.io timer:', error);
                timerData = { ok: false, data: {}, error: error.message };
                timerCompleted = true;
                if (timerTimeout) clearTimeout(timerTimeout);
                sendResponse();
              });
              
              // Set request timeout
              timerReq.setTimeout(TIMER_FETCH_TIMEOUT, () => {
                if (!timerCompleted) {
                  console.warn('[API] Timer request timeout');
                  timerReq.destroy();
                  timerData = { ok: false, data: {}, error: 'Request timeout' };
                  timerCompleted = true;
                  if (timerTimeout) clearTimeout(timerTimeout);
                  sendResponse();
                }
              });
            } else {
              // No timer_id available, try without it (gets currently highlighted timer)
              const timerUrl = `https://api.stagetimer.io/v1/get_timer?room_id=${encodeURIComponent(roomId)}&api_key=${encodeURIComponent(apiKey)}`;
              console.log('[API] No timer_id, fetching currently highlighted timer');
              const timerReq = https.get(timerUrl, (timerRes) => {
                let timerDataStr = '';
                
                if (timerRes.statusCode !== 200) {
                  console.error('[API] Stagetimer timer HTTP error:', timerRes.statusCode, timerRes.statusMessage);
                  timerRes.on('data', () => {}); // Drain response
                  timerRes.on('end', () => {
                    if (!timerCompleted) {
                      timerData = { ok: false, message: `HTTP ${timerRes.statusCode}: ${timerRes.statusMessage}`, data: {} };
                      timerCompleted = true;
                      if (timerTimeout) clearTimeout(timerTimeout);
                      sendResponse();
                    }
                  });
                  return;
                }
                
                timerRes.on('data', (chunk) => {
                  timerDataStr += chunk;
                });
                
                timerRes.on('end', () => {
                  if (timerCompleted) return; // Already handled by timeout
                  try {
                    timerData = JSON.parse(timerDataStr);
                    console.log('[API] Stagetimer timer response:', JSON.stringify(timerData, null, 2));
                    console.log('[API] Timer data.name:', timerData?.data?.name);
                    console.log('[API] Timer data.speaker:', timerData?.data?.speaker);
                    timerCompleted = true;
                    if (timerTimeout) clearTimeout(timerTimeout);
                    sendResponse();
                  } catch (error) {
                    console.error('[API] Error parsing stagetimer timer response:', error);
                    console.error('[API] Raw timer response:', timerDataStr);
                    timerData = { ok: false, data: {}, error: error.message };
                    timerCompleted = true;
                    if (timerTimeout) clearTimeout(timerTimeout);
                    sendResponse();
                  }
                });
              });
              
              timerReq.on('error', (error) => {
                if (timerCompleted) return; // Already handled by timeout
                console.error('[API] Error calling stagetimer.io timer:', error);
                timerData = { ok: false, data: {}, error: error.message };
                timerCompleted = true;
                if (timerTimeout) clearTimeout(timerTimeout);
                sendResponse();
              });
              
              // Set request timeout
              timerReq.setTimeout(TIMER_FETCH_TIMEOUT, () => {
                if (!timerCompleted) {
                  console.warn('[API] Timer request timeout');
                  timerReq.destroy();
                  timerData = { ok: false, data: {}, error: 'Request timeout' };
                  timerCompleted = true;
                  if (timerTimeout) clearTimeout(timerTimeout);
                  sendResponse();
                }
              });
            }
            
            sendResponse();
          } catch (error) {
            console.error('[API] Error parsing stagetimer status response:', error);
            statusData = { ok: false, message: 'Failed to parse status response' };
            completed++;
            sendResponse();
          }
        });
      });
      
      statusReq.on('error', (error) => {
        console.error('[API] Error calling stagetimer.io status:', error);
        statusData = { ok: false, message: 'Failed to connect: ' + error.message };
        completed++;
        sendResponse();
      });
      
      // Set request timeout for status
      statusReq.setTimeout(10000, () => {
        console.warn('[API] Status request timeout');
        statusReq.destroy();
        if (!statusData) {
          statusData = { ok: false, message: 'Request timeout' };
          completed++;
          sendResponse();
        }
      });
      
      // Fetch messages
      https.get(messagesUrl, (apiRes) => {
        let data = '';
        
        apiRes.on('data', (chunk) => {
          data += chunk;
        });
        
        apiRes.on('end', () => {
          try {
            messagesData = JSON.parse(data);
            console.log('[API] Stagetimer messages response:', JSON.stringify(messagesData, null, 2));
            completed++;
            sendResponse();
          } catch (error) {
            console.error('[API] Error parsing stagetimer messages response:', error);
            messagesData = { ok: false, data: { messages: [] } };
            completed++;
            sendResponse();
          }
        });
      }).on('error', (error) => {
        console.error('[API] Error calling stagetimer.io messages:', error);
        messagesData = { ok: false, data: { messages: [] } };
        completed++;
        sendResponse();
      });
      
      return;
    }
    
    // GET /api/presets - Get all preset presentations
    if (req.method === 'GET' && apiReqPath === '/api/presets') {
      console.log('[API] GET /api/presets - Loading presets');
      const prefs = loadPreferences();
      const urls = getPresetUrlsFromPrefs(prefs);
      res.writeHead(200, {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
      });
      res.end(JSON.stringify({ presetUrls: urls }));
      return;
    }

    // GET /api/debug/preferences - Debug endpoint for preferences file
    if (req.method === 'GET' && req.url === '/api/debug/preferences') {
      try {
        const prefsPath = getPreferencesPath();
        const prefsDir = path.dirname(prefsPath);
        const exists = fs.existsSync(prefsPath);
        const dirExists = fs.existsSync(prefsDir);
        
        let stats = null;
        let content = null;
        let dirWritable = false;
        
        if (exists) {
          stats = fs.statSync(prefsPath);
          try {
            content = fs.readFileSync(prefsPath, 'utf8');
          } catch (e) {
            content = `Error reading file: ${e.message}`;
          }
        }
        
        try {
          fs.accessSync(prefsDir, fs.constants.W_OK);
          dirWritable = true;
        } catch (e) {
          dirWritable = false;
        }
        
        res.writeHead(200, { 
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type'
        });
        res.end(JSON.stringify({
          path: prefsPath,
          directory: prefsDir,
          fileExists: exists,
          directoryExists: dirExists,
          directoryWritable: dirWritable,
          fileSize: stats ? stats.size : null,
          fileModified: stats ? stats.mtime : null,
          fileContent: content,
          preferences: sanitizePreferencesForClient(loadPreferences()),
          platform: process.platform,
          userData: app.getPath('userData')
        }));
      } catch (error) {
        res.writeHead(500, { 
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type'
        });
        res.end(JSON.stringify({ error: error.message, stack: error.stack }));
      }
      return;
    }

    // POST /api/stagetimer-settings - Save stagetimer configuration
    if (req.method === 'POST' && req.url === '/api/stagetimer-settings') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      req.on('end', () => {
        try {
          const data = JSON.parse(body);
          const prefs = loadPreferences();
          
          if (data.roomId !== undefined) {
            prefs.stagetimerRoomId = data.roomId || null;
          }
          if (data.apiKey !== undefined) {
            prefs.stagetimerApiKey = data.apiKey || null;
          }
          if (data.enabled !== undefined) {
            prefs.stagetimerEnabled = data.enabled !== false;
          }
          if (data.visible !== undefined) {
            prefs.stagetimerVisible = data.visible !== false;
          } else {
            // Default to true if not set
            prefs.stagetimerVisible = true;
          }
          
          savePreferences(prefs);
          
          res.writeHead(200, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ success: true, message: 'Stagetimer settings saved' }));
        } catch (error) {
          console.error('[API] Error saving stagetimer settings:', error);
          res.writeHead(500, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ success: false, error: error.message }));
        }
      });
      return;
    }
    
    // GET /api/stagetimer-settings - Get stagetimer configuration
    if (req.method === 'GET' && req.url === '/api/stagetimer-settings') {
      const prefs = loadPreferences();
      res.writeHead(200, { 
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      });
      res.end(JSON.stringify({
        roomId: prefs.stagetimerRoomId || '',
        apiKey: prefs.stagetimerApiKey || '',
        enabled: prefs.stagetimerEnabled !== false,
        visible: prefs.stagetimerVisible !== false && prefs.stagetimerVisible !== undefined ? prefs.stagetimerVisible : true
      }));
      return;
    }
    
    // POST /api/presets - Set preset presentations
    if (req.method === 'POST' && apiReqPath === '/api/presets') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      
      req.on('end', () => {
        try {
          logDebug('[API] POST /api/presets - Received body:', body);
          const data = JSON.parse(body);
          const prefs = loadPreferences();
          if (Array.isArray(data.presetUrls)) {
            prefs.presetUrls = data.presetUrls.map(u => String(u || '').trim()).filter(Boolean);
          }
          savePreferences(prefs);
          const verifyPrefs = loadPreferences();
          res.writeHead(200, {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type'
          });
          res.end(JSON.stringify({
            success: true,
            message: 'Presets saved',
            saved: { presetUrls: getPresetUrlsFromPrefs(verifyPrefs) }
          }));
        } catch (error) {
          console.error('[API] Error saving presets:', error);
          console.error('[API] Error stack:', error.stack);
          res.writeHead(500, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type'
          });
          res.end(JSON.stringify({ 
            error: error.message,
            code: error.code || 'UNKNOWN',
            details: process.platform === 'darwin' ? 'Check Console.app for detailed logs' : 'Check console output'
          }));
        }
      });
      return;
    }

    // POST /api/open-preset - Open a preset by index (1-based)
    if (req.method === 'POST' && apiReqPath === '/api/open-preset') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      
      req.on('end', async () => {
        try {
          const data = JSON.parse(body);
          const presetNumber = parseInt(data.preset, 10);
          const prefs = loadPreferences();
          const urls = getPresetUrlsFromPrefs(prefs);
          if (isNaN(presetNumber) || presetNumber < 1 || presetNumber > urls.length) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: `Preset must be between 1 and ${urls.length}` }));
            return;
          }
          const url = urls[presetNumber - 1];
          if (!url) {
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: `Preset ${presetNumber} is not configured` }));
            return;
          }
          
          // Forward to open-presentation endpoint logic
          // We'll reuse the same code path
          console.log(`[API] Opening preset ${presetNumber}: ${url}`);
          
          // Close any existing presentation windows
          try {
            if (notesWindow && !notesWindow.isDestroyed()) {
              console.log('[API] Closing existing notes window');
              notesWindow.removeAllListeners('closed');
              notesWindow.close();
              notesWindow = null;
            }
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Closing existing presentation window');
              presentationWindow.removeAllListeners('closed');
              presentationWindow.close();
              presentationWindow = null;
            }
            currentSlide = null;
          } catch (error) {
            console.error('[API] Error closing existing windows:', error.message);
          }
          
          // Load preferences for monitor selection
          const displays = screen.getAllDisplays();
          const presentationDisplayId = Number(prefs.presentationDisplayId);
          const notesDisplayId = Number(prefs.notesDisplayId);
          const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
          const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];
          
          console.log('[API] Using presentation display:', presentationDisplay.id);
          console.log('[API] Using notes display:', notesDisplay.id);
          
          // Create the presentation window (reuse open-presentation logic)
          // Note: Don't use fullscreen: true in constructor as it creates a new Space on macOS
          // We'll use setSimpleFullScreen() after creation to avoid Spaces conflicts
          presentationWindow = new BrowserWindow(getPresentationBrowserWindowOptions(presentationDisplay.bounds));
          
          // Set simple fullscreen on macOS to avoid Spaces conflicts
          if (process.platform === 'darwin') {
            applyPresentationFullscreenChrome(presentationWindow, prefs);
          }
          attachCrashHandlers(presentationWindow, 'presentation');
          
          // Set up window handlers (same as open-presentation; notes open at notes-display size for narrow-preview layout)
          presentationWindow.webContents.setWindowOpenHandler(({ url }) => {
            const windowOptions = getSpeakerNotesWindowOptions(notesDisplay);
            return { action: 'allow', overrideBrowserWindowOptions: windowOptions };
          });
          
          const windowCreatedListener = (event, window) => {
            if (window !== presentationWindow && window !== mainWindow) {
              notesWindow = window;
              onNotesWindowCreated(window);
              attachCrashHandlers(window, 'notes');
              window.webContents.on('before-input-event', (event, input) => {
                if (input.key === 'Escape' && input.type === 'keyDown') {
                  event.preventDefault();
                  if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
                  if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
                }
              });
              applySpeakerNotesInitialGeometry(window, notesDisplay, null);

              app.removeListener('browser-window-created', windowCreatedListener);
            }
          };
          app.on('browser-window-created', windowCreatedListener);

          let sKeyPressed = false;
          const navigationListener = async (event, navUrl) => {
            const isPresentMode = (navUrl.includes('/present/') || navUrl.includes('localpresent')) && !navUrl.includes('/presentation/');
            if (isPresentMode && !sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
              sKeyPressed = true;
              await new Promise(resolve => setTimeout(resolve, 300));
              if (presentationWindow && !presentationWindow.isDestroyed()) {
                presentationWindow.focus();
                await new Promise(resolve => setTimeout(resolve, 50));
                presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
                presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
                presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
                presentationWindow.webContents.removeListener('did-navigate', navigationListener);
              }
            }
          };
          presentationWindow.webContents.on('did-navigate', navigationListener);
          
          presentationWindow.webContents.once('did-finish-load', async () => {
            if (!presentationWindow || presentationWindow.isDestroyed()) return;
            await new Promise(resolve => setTimeout(resolve, 200));
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 50));
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
            }
          });
          
          setTimeout(async () => {
            if (!sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
              sKeyPressed = true;
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 50));
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
              presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
              if (presentationWindow && !presentationWindow.isDestroyed()) {
                presentationWindow.webContents.removeListener('did-navigate', navigationListener);
              }
            }
          }, 1000);
          
          presentationWindow.on('closed', () => {
            presentationWindow = null;
            currentSlide = null;
          });
          
          presentationWindow.webContents.on('before-input-event', (event, input) => {
            if (input.key === 'Escape' && input.type === 'keyDown') {
              event.preventDefault();
              if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
              if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
            }
          });
          
          const presentUrl = toPresentUrl(url);
          console.log('[API] Loading PRESENT URL:', presentUrl);
          lastPresentationUrl = url;
          currentSlide = 1;
          resetNotesZoomForNewPresentation();
          presentationWindow.loadURL(presentUrl);
          presentationWindow.show();
          
          // Ensure fullscreen on macOS
          presentationWindow.once('ready-to-show', () => {
            if (process.platform === 'darwin' && presentationWindow && !presentationWindow.isDestroyed()) {
              presentationWindow.setBounds({
                x: presentationDisplay.bounds.x,
                y: presentationDisplay.bounds.y,
                width: presentationDisplay.bounds.width,
                height: presentationDisplay.bounds.height
              });
              setTimeout(() => {
                if (presentationWindow && !presentationWindow.isDestroyed()) {
                  applyPresentationFullscreenChrome(presentationWindow, prefs);
                }
              }, 50);
            }
          });
          
          if (process.platform === 'darwin') {
            setTimeout(() => {
              if (presentationWindow && !presentationWindow.isDestroyed() && presentationFullscreenNeedsReapply(presentationWindow, prefs)) {
                presentationWindow.setBounds({
                  x: presentationDisplay.bounds.x,
                  y: presentationDisplay.bounds.y,
                  width: presentationDisplay.bounds.width,
                  height: presentationDisplay.bounds.height
                });
                applyPresentationFullscreenChrome(presentationWindow, prefs);
              }
            }, 200);
          }
          
          // Broadcast to backups (async, don't wait)
          sendToBackups('/api/open-preset', { preset: presetNumber }).catch(err => {
            console.error('[Backup] Error broadcasting open-preset:', err);
          });
          
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: true, message: `Preset ${presetNumber} opened`, url: url }));
        } catch (error) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        }
      });
      return;
    }

    // 404 for unknown endpoints
    res.writeHead(404, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Not found' }));
  });
  
  const prefs = loadPreferences();
  const apiPort = prefs.apiPort || DEFAULT_API_PORT;
  
  httpServer.listen(apiPort, '0.0.0.0', () => {
    console.log(`[API] HTTP server listening on http://0.0.0.0:${apiPort}`);
  }).on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`[API] Port ${apiPort} is already in use`);
      dialog.showErrorBox(
        'Port Already in Use',
        `Port ${apiPort} is already in use. Another instance of Google Slides Opener may be running.\n\nPlease quit the other instance or change the API port in settings.`
      );
      // Don't exit the app, but the server won't start
    } else {
      console.error('[API] Server error:', err);
      dialog.showErrorBox(
        'Server Error',
        `Failed to start API server: ${err.message}`
      );
    }
  });
}

// Get HTTPS credentials for Web UI: user-provided cert/key or self-signed in userData. Returns { key, cert } or null.
function getWebUiHttpsCredentials() {
  const prefs = loadPreferences();
  if (!prefs.webUiUseHttps) return null;
  const certPath = (prefs.webUiCertPath || '').trim();
  const keyPath = (prefs.webUiKeyPath || '').trim();
  if (certPath && keyPath && fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    try {
      return {
        key: fs.readFileSync(keyPath, 'utf8'),
        cert: fs.readFileSync(certPath, 'utf8')
      };
    } catch (e) {
      console.error('[Web UI] Failed to read HTTPS cert/key files:', e.message);
      return null;
    }
  }
  const userData = app.getPath('userData');
  const selfKeyPath = path.join(userData, 'webui-selfsigned-key.pem');
  const selfCertPath = path.join(userData, 'webui-selfsigned-cert.pem');
  if (fs.existsSync(selfKeyPath) && fs.existsSync(selfCertPath)) {
    try {
      return {
        key: fs.readFileSync(selfKeyPath, 'utf8'),
        cert: fs.readFileSync(selfCertPath, 'utf8')
      };
    } catch (e) {
      console.error('[Web UI] Failed to read self-signed cert:', e.message);
      return null;
    }
  }
  try {
    execSync(
      `openssl req -x509 -newkey rsa:2048 -keyout "${selfKeyPath}" -out "${selfCertPath}" -days 365 -nodes -subj "/CN=localhost"`,
      { stdio: 'pipe' }
    );
    return {
      key: fs.readFileSync(selfKeyPath, 'utf8'),
      cert: fs.readFileSync(selfCertPath, 'utf8')
    };
  } catch (e) {
    console.error('[Web UI] Could not generate self-signed certificate (openssl not available?):', e.message);
    return null;
  }
}

function getWebUiLocalOriginForTunnel() {
  const creds = getWebUiHttpsCredentials();
  const prefs = loadPreferences();
  let webUiPort = prefs.webUiPort || DEFAULT_WEB_UI_PORT;
  if (creds && webUiPort === 80) {
    webUiPort = DEFAULT_WEB_UI_HTTPS_PORT;
  }
  const protocol = creds ? 'https' : 'http';
  return `${protocol}://127.0.0.1:${webUiPort}`;
}

function getCloudflaredBinaryPath() {
  const base = app.isPackaged
    ? process.resourcesPath
    : path.join(__dirname, 'resources');
  if (process.platform === 'win32') {
    return path.join(base, 'cloudflared', 'cloudflared-windows-amd64.exe');
  }
  const arch = process.arch === 'arm64' ? 'arm64' : 'amd64';
  if (process.platform === 'darwin') {
    return path.join(base, 'cloudflared', `cloudflared-darwin-${arch}`);
  }
  return path.join(base, 'cloudflared', `cloudflared-linux-${arch}`);
}

function extractTrycloudflareUrl(chunk) {
  const text = chunk.toString();
  const strict = /https:\/\/[a-z0-9-]+\.trycloudflare\.com\b/;
  const loose = /https:\/\/[^\s"'<>]+\.trycloudflare\.com\b/i;
  let m = text.match(strict);
  if (m) return m[0];
  m = text.match(loose);
  return m ? m[0] : null;
}

function broadcastTunnelUrl(url) {
  try {
    BrowserWindow.getAllWindows().forEach((w) => {
      if (!w.isDestroyed()) {
        w.webContents.send('tunnel-url-changed', url);
      }
    });
  } catch (e) {
    logDebug('[Tunnel] broadcastTunnelUrl:', e);
  }
}

async function showTunnelQrOverlay(url, durationMs) {
  hideTunnelQrOverlay();

  const prefs = loadPreferences();
  const displays = screen.getAllDisplays();
  const notesDisplay = displays.find(d => d.id === Number(prefs.notesDisplayId)) || displays[0];
  const b = notesDisplay.bounds;

  const QR_PX = 280;
  const PAD = 20;
  const W = QR_PX + PAD * 2;
  const H = QR_PX + PAD * 2;
  const x = b.x + Math.floor((b.width - W) / 2);
  const y = b.y + Math.floor((b.height - H) / 2);

  tunnelQrWindow = new BrowserWindow({
    x, y, width: W, height: H,
    frame: false,
    transparent: true,
    alwaysOnTop: true,
    skipTaskbar: true,
    resizable: false,
    focusable: false,
    webPreferences: { nodeIntegration: false, contextIsolation: true }
  });

  const qrDataUrl = await QRCode.toDataURL(url, { width: QR_PX, margin: 1 });

  const html = `<!DOCTYPE html><html><body style="margin:0;background:rgba(0,0,0,0.82);border-radius:12px;display:flex;align-items:center;justify-content:center;box-sizing:border-box;width:100%;height:100vh;">
    <img src="${qrDataUrl}" alt="" style="width:${QR_PX}px;height:${QR_PX}px;border-radius:8px;display:block;" />
  </body></html>`;

  tunnelQrWindow.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
  tunnelQrWindow.on('closed', () => { tunnelQrWindow = null; });

  tunnelQrHideTimer = setTimeout(hideTunnelQrOverlay, durationMs);
}

function hideTunnelQrOverlay() {
  if (tunnelQrHideTimer) { clearTimeout(tunnelQrHideTimer); tunnelQrHideTimer = null; }
  if (tunnelQrWindow && !tunnelQrWindow.isDestroyed()) { tunnelQrWindow.close(); }
  tunnelQrWindow = null;
}

function stopCloudflaredTunnel() {
  hideTunnelQrOverlay();
  if (cloudflaredKillTimer) {
    clearTimeout(cloudflaredKillTimer);
    cloudflaredKillTimer = null;
  }
  tunnelUrl = null;
  if (!cloudflaredProcess) {
    broadcastTunnelUrl(null);
    return;
  }
  const proc = cloudflaredProcess;
  cloudflaredProcess = null;
  try {
    proc.kill('SIGTERM');
  } catch (e) {
    // ignore
  }
  cloudflaredKillTimer = setTimeout(() => {
    cloudflaredKillTimer = null;
    try {
      if (proc && !proc.killed) {
        proc.kill('SIGKILL');
      }
    } catch (e) {
      // ignore
    }
  }, 5000);
  broadcastTunnelUrl(null);
}

function startCloudflaredTunnel() {
  const prefs = loadPreferences();
  if (!prefs.cloudflaredEnabled) return;

  const bin = getCloudflaredBinaryPath();
  if (!fs.existsSync(bin)) {
    if (app.isPackaged) {
      logError(
        '[Tunnel] cloudflared binary missing:',
        bin,
        '— The app was built without bundled cloudflared. Fix: (1) On a machine with the source repo, run `yarn download:cloudflared` then rebuild the .app; or (2) Download the matching cloudflared release for your Mac CPU, then place the binary at this path inside the .app (see docs). `yarn download:cloudflared` only works in the project root (where package.json is), not from Applications.'
      );
    } else {
      logError('[Tunnel] cloudflared binary missing:', bin, '(from repo root run: yarn download:cloudflared)');
    }
    broadcastTunnelUrl(null);
    return;
  }

  if (cloudflaredProcess) return;

  const origin = getWebUiLocalOriginForTunnel();
  // Web UI often uses a self-signed cert; cloudflared rejects it without this → edge returns 502
  const tunnelArgs = ['tunnel', '--url', origin];
  if (origin.startsWith('https://')) {
    tunnelArgs.push('--no-tls-verify');
  }

  try {
    cloudflaredProcess = spawn(bin, tunnelArgs, {
      stdio: ['ignore', 'pipe', 'pipe']
    });
  } catch (err) {
    logError('[Tunnel] Failed to spawn cloudflared:', err.message);
    cloudflaredProcess = null;
    return;
  }

  const onData = (data) => {
    const found = extractTrycloudflareUrl(data);
    if (found && !tunnelUrl) {
      tunnelUrl = found;
      logInfo('[Tunnel] URL:', tunnelUrl);
      broadcastTunnelUrl(tunnelUrl);
    }
    logDebug('[Tunnel]', data.toString().slice(0, 240));
  };

  cloudflaredProcess.stdout.on('data', onData);
  cloudflaredProcess.stderr.on('data', onData);

  cloudflaredProcess.on('exit', (code) => {
    logInfo('[Tunnel] cloudflared exited, code:', code);
    if (cloudflaredKillTimer) {
      clearTimeout(cloudflaredKillTimer);
      cloudflaredKillTimer = null;
    }
    cloudflaredProcess = null;
    tunnelUrl = null;
    broadcastTunnelUrl(null);
  });
}

// Start web UI server for preset management
function startWebUiServer() {
  const requestHandler = (req, res) => {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    
    if (req.method === 'OPTIONS') {
      res.writeHead(200);
      res.end();
      return;
    }

    // Controller allowlist: restrict access to the Web UI (and its /api proxy)
    try {
      const prefs = loadPreferences();
      if (!isControllerAllowedRequest(req, prefs)) {
        res.writeHead(403, { 'Content-Type': 'text/plain' });
        res.end('Forbidden');
        return;
      }
    } catch (e) {
      // default allow if something goes wrong
    }

    const reqPath = String(req.url || '').split('?')[0];

    // Tunnel clients (localhost via cloudflared): optional PIN gate before Web UI + /api proxy
    try {
      const prefsGate = loadPreferences();
      if (isWebUiPinGateActiveForRequest(req, prefsGate) && isWebUiTunnelPinConfigured(prefsGate)) {
        const tunnelAuthed = isValidTunnelWebUiSessionCookie(readWebUiTunnelCookie(req), prefsGate);
        if (!tunnelAuthed) {
          if (reqPath === '/tunnel-unlock' && req.method === 'GET') {
            res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
            res.end(buildTunnelUnlockHtml());
            return;
          }
          if (reqPath === '/tunnel-unlock' && req.method === 'POST') {
            let body = '';
            req.on('data', (chunk) => {
              body += chunk.toString();
              if (body.length > TUNNEL_PIN_UNLOCK_MAX_BODY) {
                try {
                  req.destroy();
                } catch (e) {
                  // ignore
                }
              }
            });
            req.on('end', () => {
              try {
                const key = tunnelUnlockClientKey(req);
                if (isTunnelPinUnlockBlocked(key)) {
                  res.writeHead(429, {
                    'Content-Type': 'application/json',
                    'Access-Control-Allow-Origin': '*',
                    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
                    'Access-Control-Allow-Headers': 'Content-Type'
                  });
                  res.end(JSON.stringify({ success: false, error: 'Too many attempts. Try again later.' }));
                  return;
                }
                const j = JSON.parse(body || '{}');
                const pin = j.pin;
                if (!verifyWebUiTunnelPin(pin, prefsGate)) {
                  recordTunnelPinFailure(key);
                  res.writeHead(401, {
                    'Content-Type': 'application/json',
                    'Access-Control-Allow-Origin': '*',
                    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
                    'Access-Control-Allow-Headers': 'Content-Type'
                  });
                  res.end(JSON.stringify({ success: false, error: 'Invalid PIN' }));
                  return;
                }
                clearTunnelPinFailures(key);
                res.writeHead(200, {
                  'Content-Type': 'application/json',
                  'Access-Control-Allow-Origin': '*',
                  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
                  'Access-Control-Allow-Headers': 'Content-Type',
                  'Set-Cookie': buildSetTunnelSessionCookieHeader(prefsGate, req)
                });
                res.end(JSON.stringify({ success: true }));
              } catch (e) {
                res.writeHead(400, {
                  'Content-Type': 'application/json',
                  'Access-Control-Allow-Origin': '*',
                  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
                  'Access-Control-Allow-Headers': 'Content-Type'
                });
                res.end(JSON.stringify({ success: false, error: 'Bad request' }));
              }
            });
            return;
          }
          if (reqPath.startsWith('/api/')) {
            res.writeHead(401, {
              'Content-Type': 'application/json',
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
              'Access-Control-Allow-Headers': 'Content-Type'
            });
            res.end(JSON.stringify({
              success: false,
              error: 'Web UI PIN required. Open /tunnel-unlock in this browser, then try again.'
            }));
            return;
          }
          if (req.method === 'GET') {
            res.writeHead(302, { Location: '/tunnel-unlock', 'Cache-Control': 'no-store' });
            res.end();
            return;
          }
          res.writeHead(403, { 'Content-Type': 'text/plain' });
          res.end('Forbidden');
          return;
        }
      }
    } catch (e) {
      logDebug('[Web UI] tunnel pin gate:', e);
    }

    // GET /custom-style.css - Serve user-selected CSS override (white-label)
    if (req.method === 'GET' && reqPath === '/custom-style.css') {
      try {
        const prefs = loadPreferences();
        const cssPath = prefs.webUiCustomCssPath;
        if (!cssPath || !fs.existsSync(cssPath)) {
          res.writeHead(404, { 'Content-Type': 'text/plain' });
          res.end('Not found');
          return;
        }
        const css = fs.readFileSync(cssPath, 'utf8');
        res.writeHead(200, { 'Content-Type': 'text/css; charset=utf-8', 'Cache-Control': 'no-cache' });
        res.end(css);
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end('Error loading custom CSS');
      }
      return;
    }

    // GET /custom-logo - Serve user-selected brand logo (light/dark themes only)
    if (req.method === 'GET' && reqPath === '/custom-logo') {
      try {
        const prefs = loadPreferences();
        const logoPath = prefs.webUiLogoPath;
        if (!logoPath || !fs.existsSync(logoPath)) {
          res.writeHead(404, { 'Content-Type': 'text/plain' });
          res.end('Not found');
          return;
        }
        const ext = path.extname(logoPath).toLowerCase();
        const mime = ext === '.png' ? 'image/png' : ext === '.jpg' || ext === '.jpeg' ? 'image/jpeg' : ext === '.svg' ? 'image/svg+xml' : 'image/png';
        const buf = fs.readFileSync(logoPath);
        res.writeHead(200, { 'Content-Type': mime, 'Cache-Control': 'no-cache' });
        res.end(buf);
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end('Error loading logo');
      }
      return;
    }

    // Serve favicon (prevents browser 404 spam)
    if (req.method === 'GET' && (reqPath === '/favicon.ico' || reqPath === '/favicon.png')) {
      const png = getFaviconPngBuffer();
      if (!png) {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        res.end('Not found');
        return;
      }
      res.writeHead(200, {
        'Content-Type': 'image/png',
        // Browsers cache favicons aggressively; keep it short and allow refresh.
        'Cache-Control': 'public, max-age=300'
      });
      res.end(png);
      return;
    }
    
    // GET / - Serve the web UI (use path only — req.url includes ?query and would not match '/')
    if (req.method === 'GET' && reqPath === '/') {
      // Get configured API port for the web UI
      const prefs = loadPreferences();
      const apiPort = prefs.apiPort || DEFAULT_API_PORT;
      const webUiPort = prefs.webUiPort || DEFAULT_WEB_UI_PORT;
      const webUiDebugConsoleEnabled = prefs.webUiDebugConsoleEnabled === true;
      const hasFavicon = !!getFaviconPngBuffer();
      const faviconHref = `/favicon.png?v=${encodeURIComponent(appBuildInfo.buildNumber || '0')}`;
      const webUiTheme = ['original', 'light', 'dark', 'max', 'touch', 'thumb'].includes(prefs.webUiTheme) ? prefs.webUiTheme : 'original';
      const webUiCustomCssPath = prefs.webUiCustomCssPath || '';
      const webUiLogoPath = prefs.webUiLogoPath || '';
      const showLogo = webUiTheme !== 'max' && webUiLogoPath && fs.existsSync(webUiLogoPath);
      const webUiRestrictedTunnelClient = isWebUiRestrictedTunnelClient(req, prefs);
      
      // Get version and build number
      const versionString = `v${appBuildInfo.version}.${appBuildInfo.buildNumber}`;
      
      // Get machine name or fallback to hostname
      // Escape HTML to prevent XSS
      const machineName = (prefs.machineName || os.hostname())
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');

      const html = `<!DOCTYPE html>
<html lang="en" data-gso-build="${String(appBuildInfo.buildNumber || '0')}">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Google Slides Opener - Preset Manager</title>
  ${hasFavicon ? `<link rel="icon" type="image/png" href="${faviconHref}"><link rel="shortcut icon" href="${faviconHref}">` : ``}
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Lora:ital,wght@0,400;0,500;1,400&display=swap" rel="stylesheet">
  <script src="https://cdn.socket.io/4.5.4/socket.io.min.js"></script>
  <style>
    :root {
      --faire-font-sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      --faire-font-serif: 'Lora', Georgia, 'Times New Roman', serif;
      --faire-font-mono: ui-monospace, 'SF Mono', Menlo, Consolas, monospace;
      --faire-text: #333333;
      --faire-sub: #757575;
      --faire-muted: #b5a998;
      --faire-border: #dfe0e1;
      --faire-surface: #ffffff;
      --faire-warm: #fbf8f6;
      --faire-page: #fafaf8;
      --tmr-idle-bg: #fbf8f6;
      --tmr-idle-bd: #dfe0e1;
      --tmr-idle-fg: #757575;
      --tmr-run-bg: #eef2ed;
      --tmr-run-bd: #c8d4c8;
      --tmr-run-fg: #49694c;
      --tmr-warn-bg: #f6efdb;
      --tmr-warn-bd: #d1b985;
      --tmr-warn-fg: #907c3a;
      --tmr-crit-bg: #f5dcd6;
      --tmr-crit-bd: #d9a79a;
      --tmr-crit-fg: #921100;
      --tmr-over-bg: #3a1510;
      --tmr-over-bd: #6e1100;
      --tmr-over-fg: #ffd3c9;
      /* Timer digit color per state (body copy still uses --tmr-*-fg on container) */
      --tmr-idle-clk: #333333;
      --tmr-run-clk: #2d4a30;
      --tmr-warn-clk: #5c4e1e;
      --tmr-crit-clk: #6e1100;
      --tmr-over-clk: #ffffff;
      --faire-radius: 4px;
      --faire-settings-card-radius: 10px;
    }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    html { overflow-x: hidden; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
      overflow-x: hidden;
      width: 100%;
    }
    @media (max-width: 768px) {
      body {
        padding: 12px;
        align-items: flex-start;
      }
    }
    .container {
      background: white;
      border-radius: 16px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3);
      max-width: 600px;
      width: 100%;
      min-width: 0;
      padding: 40px;
      transition: all 0.3s;
      margin-left: auto;
      margin-right: auto;
    }
    @media (max-width: 768px) {
      .container {
        padding: 16px 20px;
        width: min(100%, calc(100vw - 24px));
        max-width: calc(100vw - 24px);
      }
    }
    body.notes-visible .container,
    body.previews-visible .container {
      max-width: 85%;
    }
    @media (max-width: 768px) {
      body.notes-visible .container,
      body.previews-visible .container {
        width: min(100%, calc(100vw - 24px));
        max-width: calc(100vw - 24px);
      }
    }
    body.notes-visible .container,
    body.previews-visible .container {
      padding: 24px 28px;
    }
    @media (max-width: 768px) {
      body.notes-visible .container,
      body.previews-visible .container {
        padding: 12px 14px;
      }
    }
    h1 {
      color: #333;
      margin-top: 0;
      margin-bottom: 0;
      padding-top: 8px;
      padding-bottom: 8px;
      font-size: 28px;
      transition: all 0.3s;
      display: flex;
      align-items: center;
      gap: 12px;
    }
    body.notes-visible h1,
    body.previews-visible h1 {
      font-size: 20px;
      padding-top: 4px;
      padding-bottom: 4px;
    }
    .system-icon {
      width: 32px;
      height: 32px;
      color: #667eea;
      flex-shrink: 0;
    }
    body.notes-visible .system-icon,
    body.previews-visible .system-icon {
      width: 24px;
      height: 24px;
    }
    .web-ui-header {
      display: flex;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }
    .web-ui-brand-logo {
      max-height: 56px;
      width: auto;
      object-fit: contain;
      flex-shrink: 0;
    }
    body.notes-visible .web-ui-brand-logo,
    body.previews-visible .web-ui-brand-logo {
      max-height: 40px;
    }
    body.theme-max .web-ui-brand-logo {
      max-height: 24px;
    }
    .preset-group {
      margin-bottom: 24px;
    }
    label {
      display: block;
      font-weight: 600;
      color: #333;
      margin-bottom: 8px;
      font-size: 14px;
    }
    input[type="text"] {
      width: 100%;
      padding: 12px 16px;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      font-size: 14px;
      transition: border-color 0.2s;
    }
    input[type="text"]:focus {
      outline: none;
      border-color: #667eea;
    }
    .btn {
      width: 100%;
      padding: 14px;
      background: #667eea;
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.2s;
      margin-top: 8px;
    }
    .btn:hover {
      background: #5568d3;
    }
    .btn:active {
      transform: scale(0.98);
    }
    .btn-secondary {
      background: #6c757d;
      margin-top: 12px;
    }
    .btn-secondary:hover {
      background: #5a6268;
    }
    .status {
      position: fixed;
      bottom: 24px;
      left: 50%;
      transform: translateX(-50%) translateY(100px);
      padding: 12px 20px;
      border-radius: 8px;
      text-align: center;
      font-size: 14px;
      font-weight: 500;
      opacity: 0;
      pointer-events: none;
      z-index: 1000;
      min-width: 200px;
      max-width: 90%;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      transition: all 0.3s ease;
    }
    .status.success {
      background: #d4edda;
      color: #155724;
      border: 1px solid #c3e6cb;
      opacity: 1;
      transform: translateX(-50%) translateY(0);
      pointer-events: auto;
    }
    .status.error {
      background: #f8d7da;
      color: #721c24;
      border: 1px solid #f5c6cb;
      opacity: 1;
      transform: translateX(-50%) translateY(0);
      pointer-events: auto;
    }
    .info {
      background: #e7f3ff;
      border: 1px solid #b3d9ff;
      color: #004085;
      padding: 12px;
      border-radius: 8px;
      margin-bottom: 24px;
      font-size: 13px;
      line-height: 1.5;
    }
    .controls-section {
      margin-bottom: 30px;
      padding-top: 20px;
      border-top: 2px solid #e0e0e0;
    }
    .controls-section h3 {
      color: #333;
      font-size: 18px;
      margin-bottom: 12px;
      margin-top: 20px;
    }
    .controls-section h3:first-child {
      margin-top: 0;
    }
    .controls-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 10px;
      margin-bottom: 20px;
    }
    .web-preset-launch-row {
      display: flex;
      gap: 10px;
      margin-bottom: 10px;
      align-items: stretch;
    }
    .web-preset-launch-label {
      font-weight: 600;
      color: #333;
      padding: 12px 0;
      min-width: 120px;
      font-size: 14px;
    }
    .web-preset-launch-actions {
      display: flex;
      gap: 10px;
      flex: 1;
      min-width: 0;
    }
    .web-preset-empty {
      color: #999;
      font-style: italic;
      padding: 20px;
      text-align: center;
    }
    .web-preset-empty-link {
      display: inline;
      margin: 0;
      padding: 0;
      border: none;
      background: none;
      font: inherit;
      font-style: normal;
      font-weight: 600;
      color: #667eea;
      text-decoration: underline;
      cursor: pointer;
    }
    .web-preset-empty-link:hover {
      opacity: 0.88;
    }
    .web-ui-callout--warning {
      margin-top: 10px;
      padding: 10px 12px;
      background: #ff9800;
      color: white;
      border-radius: 4px;
      font-size: 12px;
      line-height: 1.45;
    }
    .btn-control {
      padding: 12px 16px;
      background: #f8f9fa;
      color: #333;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
    }
    .btn-control:hover {
      background: #667eea;
      color: white;
      border-color: #667eea;
      transform: translateY(-2px);
      box-shadow: 0 4px 8px rgba(102, 126, 234, 0.3);
    }
    .btn-control:active {
      transform: translateY(0);
    }
    .btn-control:disabled {
      opacity: 0.5;
      cursor: not-allowed;
      transform: none;
    }
    .btn-icon {
      width: 18px;
      height: 18px;
      stroke-width: 2.5;
    }
    .tabs {
      display: flex;
      gap: 8px;
      margin-bottom: 24px;
      border-bottom: 2px solid #e0e0e0;
      transition: all 0.3s;
    }
    body.notes-visible .tabs,
    body.previews-visible .tabs {
      margin-bottom: 12px;
      border-bottom-width: 1px;
    }
    .tab-btn {
      padding: 12px 24px;
      background: transparent;
      border: none;
      border-bottom: 3px solid transparent;
      font-size: 16px;
      font-weight: 600;
      color: #666;
      cursor: pointer;
      transition: all 0.3s;
      margin-bottom: -2px;
    }
    body.notes-visible .tab-btn,
    body.previews-visible .tab-btn {
      padding: 8px 16px;
      font-size: 13px;
      border-bottom-width: 2px;
    }
    .tab-btn:hover {
      color: #333;
      background: #f8f9fa;
    }
    .tab-btn.active {
      color: #667eea;
      border-bottom-color: #667eea;
    }
    .tab-content {
      display: none;
    }
    .tab-content.active {
      display: block;
    }
    /* Inactive panels must stay hidden: theme-max / theme-thumb used to set display:flex on #tab-remote without .active, which overrode .tab-content { display:none } and leaked Remote onto Controls/Settings (iPad). */
    #tab-remote.tab-content:not(.active),
    #tab-controls.tab-content:not(.active),
    #tab-settings.tab-content:not(.active) {
      display: none !important;
    }
    /* Floating tooltips that don't affect layout */
    .btn-control[data-tooltip] {
      position: relative;
    }
    .btn-control[data-tooltip]:hover::after {
      content: attr(data-tooltip);
      position: absolute;
      bottom: calc(100% + 8px);
      left: 50%;
      transform: translateX(-50%);
      padding: 6px 10px;
      background: #333;
      color: white;
      font-size: 12px;
      font-weight: normal;
      white-space: nowrap;
      border-radius: 4px;
      pointer-events: none;
      z-index: 1000;
      box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    }
    .btn-control[data-tooltip]:hover::before {
      content: '';
      position: absolute;
      bottom: calc(100% + 2px);
      left: 50%;
      transform: translateX(-50%);
      border: 5px solid transparent;
      border-top-color: #333;
      pointer-events: none;
      z-index: 1001;
    }
    /* Build number display */
    .build-number {
      position: fixed;
      bottom: 8px;
      left: 8px;
      font-size: 11px;
      color: #999;
      opacity: 0.7;
      z-index: 10;
    }
    /* Remote tab - big buttons for mobile */
    .remote-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 20px;
      padding-bottom: 12px;
      border-bottom: 2px solid #e0e0e0;
    }
    body.notes-visible .remote-header,
    body.previews-visible .remote-header {
      margin-bottom: 12px;
      padding-bottom: 8px;
      border-bottom-width: 1px;
    }
    .remote-header h2 {
      margin: 0;
      font-size: 20px;
      color: #333;
    }
    body.notes-visible .remote-header h2,
    body.previews-visible .remote-header h2 {
      font-size: 16px;
    }
    .notes-toggle-btn {
      background: #667eea;
      color: white;
      border: none;
      border-radius: 8px;
      padding: 10px 14px;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 14px;
      transition: all 0.2s;
    }
    .preview-toggle-btn {
      background: #667eea;
      color: white;
      border: none;
      border-radius: 8px;
      padding: 10px 14px;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 14px;
      transition: all 0.2s;
    }
    .notes-toggle-btn:hover {
      background: #5568d3;
    }
    .preview-toggle-btn:hover {
      background: #5568d3;
    }
    .notes-toggle-btn.active {
      background: #764ba2;
    }
    .preview-toggle-btn.active {
      background: #764ba2;
    }
    .notes-toggle-btn svg {
      width: 18px;
      height: 18px;
    }
    .preview-toggle-btn svg {
      width: 18px;
      height: 18px;
    }
        .preview-toggle-btn:disabled,
        .preview-toggle-btn.btn-disabled {
          opacity: 0.45;
          cursor: not-allowed;
        }
    .remote-controls {
      display: flex;
      flex-direction: row;
      gap: 20px;
      padding: 20px 0;
      transition: all 0.3s;
    }
    .remote-controls.with-notes {
      gap: 20px;
    }
    .remote-controls.with-panel {
      gap: 20px;
    }
    .remote-btn {
      flex: 0 0 calc(50% - 10px);
      padding: 40px 20px;
      font-size: 24px;
      font-weight: 700;
      border-radius: 12px;
      border: none;
      cursor: pointer;
      transition: all 0.3s;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 12px;
      min-height: 120px;
    }
    .remote-controls.with-notes .remote-btn,
    .remote-controls.with-panel .remote-btn {
      padding: 20px 16px;
      font-size: 18px;
      min-height: 70px;
    }
    .remote-btn-prev {
      background: #667eea;
      color: white;
    }
    .remote-btn-prev:hover {
      background: #5568d3;
      transform: scale(1.02);
    }
    .remote-btn-next {
      background: #667eea;
      color: white;
    }
    .remote-btn-next:hover {
      background: #5568d3;
      transform: scale(1.02);
    }
    .remote-btn:active {
      transform: scale(0.98);
    }
    .remote-btn svg {
      width: 32px;
      height: 32px;
      transition: all 0.3s;
    }
    .remote-controls.with-notes .remote-btn svg,
    .remote-controls.with-panel .remote-btn svg {
      width: 24px;
      height: 24px;
    }
    /* Speaker notes display */
    .speaker-notes-container {
      display: none;
      margin-top: 12px;
      transition: all 0.3s;
    }
    .speaker-notes-container.visible {
      display: block;
    }

    /* Slide previews display (current + next) */
    .slide-previews-container {
      display: none;
      margin-top: 12px;
      transition: all 0.3s;
    }
    .slide-previews-container.visible {
      display: block;
    }
    .slide-previews-grid {
      background: #f8f9fa;
      border: 2px solid #e0e0e0;
      border-radius: 12px;
      padding: 14px;
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      align-items: flex-start;
      justify-content: center;
    }
    .slide-preview-card {
      display: flex;
      flex-direction: column;
      gap: 8px;
      align-items: center;
      width: min(220px, 45vw);
    }
    .slide-preview-card.clickable {
      cursor: pointer;
      border-radius: 12px;
      padding: 4px;
      margin: -4px;
      transition: background 0.15s ease, transform 0.1s ease;
    }
    .slide-preview-card.clickable:hover {
      background: rgba(0, 0, 0, 0.06);
    }
    .slide-preview-card.clickable:active {
      transform: scale(0.98);
    }
    .slide-preview-label {
      font-size: 12px;
      font-weight: 700;
      color: #444;
      text-align: center;
    }
    .slide-preview-img {
      width: 100%;
      max-width: 200px;
      height: auto;
      max-height: 200px;
      border-radius: 10px;
      border: 1px solid #ddd;
      background: white;
      object-fit: contain;
    }
    .slide-preview-img.empty {
      opacity: 0.18;
      background:
        linear-gradient(45deg, rgba(0,0,0,0.06) 25%, transparent 25%, transparent 75%, rgba(0,0,0,0.06) 75%, rgba(0,0,0,0.06)),
        linear-gradient(45deg, rgba(0,0,0,0.06) 25%, transparent 25%, transparent 75%, rgba(0,0,0,0.06) 75%, rgba(0,0,0,0.06));
      background-position: 0 0, 10px 10px;
      background-size: 20px 20px;
    }
    .speaker-notes-content-wrapper {
      background: #f8f9fa;
      border: 2px solid #e0e0e0;
      border-radius: 12px;
      padding: 16px;
      height: 400px;
      overflow-y: auto;
      overflow-x: hidden;
    }
    .speaker-notes-content {
      color: #333;
      font-size: 18px;
      line-height: 1.6;
      white-space: pre-wrap;
      word-wrap: break-word;
    }
    .speaker-notes-content.zoom-small {
      font-size: 16px;
    }
    .speaker-notes-content.zoom-large {
      font-size: 22px;
    }
    .notes-zoom-controls {
      display: none;
      justify-content: center;
      gap: 10px;
      margin-top: 12px;
      position: sticky;
      bottom: 0;
      background: white;
      padding: 8px 0;
      z-index: 10;
    }
    .notes-zoom-controls.visible {
      display: flex;
    }
    .notes-zoom-btn {
      background: #f8f9fa;
      border: 2px solid #e0e0e0;
      border-radius: 6px;
      padding: 8px 16px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      color: #333;
      transition: all 0.2s;
    }
    .notes-zoom-btn:hover {
      background: #667eea;
      color: white;
      border-color: #667eea;
    }
    /* Bigger slide control buttons */
    .btn-control-large {
      padding: 20px 24px;
      font-size: 18px;
      min-height: 60px;
    }
    /* Stagetimer display */
    .stagetimer-container {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      border-radius: 12px;
      padding: 16px 20px;
      padding-bottom: 16px;
      margin-bottom: 20px;
      text-align: center;
      color: white;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      position: relative;
      height: 160px;
      overflow: visible;
    }
    .stagetimer-container.error {
      background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%);
    }
    .stagetimer-container.disabled {
      background: #e0e0e0;
      color: #666;
    }
    .stagetimer-label {
      font-size: 16px;
      opacity: 0.95;
      margin-bottom: 6px;
      font-weight: 600;
    }
    .stagetimer-time {
      font-size: 42px;
      font-weight: 700;
      font-variant-numeric: tabular-nums;
      letter-spacing: 2px;
      margin: 4px 0;
      line-height: 1.2;
    }
    .stagetimer-status {
      font-size: 14px;
      opacity: 0.9;
      margin-top: 6px;
      font-weight: 500;
    }
    .stagetimer-name {
      display: none;
    }
    .stagetimer-container.running {
      background: linear-gradient(135deg, #4caf50 0%, #388e3c 100%);
    }
    .stagetimer-container.warning {
      background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%);
    }
    .stagetimer-container.critical {
      background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%);
    }
    /* Stagetimer messages - absolutely positioned to prevent layout shift */
    .stagetimer-messages {
      position: absolute;
      bottom: 0;
      left: 0;
      right: 0;
      margin: 0;
      padding: 12px 24px 16px 24px;
      border-top: 1px solid rgba(255, 255, 255, 0.3);
      background: linear-gradient(to top, rgba(0, 0, 0, 0.4) 0%, rgba(0, 0, 0, 0.2) 50%, transparent 100%);
      border-radius: 0 0 12px 12px;
      max-height: 100px;
      overflow-y: auto;
      overflow-x: hidden;
      display: none;
      backdrop-filter: blur(8px);
      z-index: 10;
    }
    .stagetimer-messages.visible {
      display: block;
    }
    .stagetimer-message {
      background: rgba(255, 255, 255, 0.15);
      border-radius: 8px;
      padding: 12px 16px;
      margin-bottom: 8px;
      font-size: 14px;
      line-height: 1.5;
      backdrop-filter: blur(10px);
    }
    .stagetimer-message:last-child {
      margin-bottom: 0;
    }
    .stagetimer-message.white {
      background: rgba(255, 255, 255, 0.2);
      color: white;
    }
    .stagetimer-message.green {
      background: rgba(76, 175, 80, 0.3);
      color: #c8e6c9;
    }
    .stagetimer-message.red {
      background: rgba(244, 67, 54, 0.3);
      color: #ffcdd2;
    }
    .stagetimer-message.bold {
      font-weight: 700;
    }
    .stagetimer-message.uppercase {
      text-transform: uppercase;
    }
    .stagetimer-row {
      display: flex;
      align-items: center;
      gap: 10px;
      width: 100%;
    }
    .stagetimer-dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: currentColor;
      flex-shrink: 0;
    }
    .stagetimer-container .stagetimer-label { flex: 1; min-width: 0; text-align: left; }
    .stagetimer-container .stagetimer-time { flex-shrink: 0; margin-left: auto; }
    .remote-header-compact {
      display: flex;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
      padding: 8px 0 12px;
    }
    .remote-status-dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: #43a047;
      flex-shrink: 0;
    }
    .remote-machine-name {
      font-size: 14px;
      font-weight: 600;
      color: #333;
      flex: 1;
      min-width: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .slide-counter {
      font-variant-numeric: tabular-nums;
      color: #757575;
      font-size: 12px;
      margin-left: auto;
      padding-right: 4px;
    }
    .speaker-notes-toolbar-label {
      flex: 1;
      font-size: 10.5px;
      letter-spacing: 0.7px;
      text-transform: uppercase;
      color: #757575;
      min-width: 0;
    }
    .notes-zoom-toolbar-actions {
      display: flex;
      align-items: center;
      gap: 6px;
      flex-shrink: 0;
      margin-left: auto;
    }
    .notes-zoom-controls {
      flex-wrap: wrap;
    }
    .bottom-tabs {
      display: none;
    }
    .remote-btn .remote-btn-label {
      display: block;
      margin-top: 4px;
      font-size: 11px;
      letter-spacing: 0.5px;
      text-transform: uppercase;
      font-weight: 600;
    }
    .toggle-btn-text {
      margin-left: 6px;
    }
  </style>
  <style id="theme-overrides">
    body:not(.theme-light) .remote-machine-name,
    body:not(.theme-light) .remote-status-dot { display: none; }
    /* Light: V2-C Faire layout (full-height remote, bottom tabs) */
    body.theme-light {
      /* Phone-first column width; widened on tablet/desktop (see @media min-width below) */
      --tl-band: min(440px, 100vw);
      background: var(--faire-page);
      /* Override base body padding: 20px — otherwise height:100vh .container + padding clips bottom nav */
      padding: 0;
      min-height: 100vh;
      min-height: 100dvh;
      font-family: var(--faire-font-sans);
      color: var(--faire-text);
      align-items: stretch;
      justify-content: flex-start;
    }
    @media (min-width: 769px) {
      body.theme-light { --tl-band: min(1000px, calc(100vw - 48px)); }
      body.theme-light.notes-visible,
      body.theme-light.previews-visible { --tl-band: min(1280px, calc(100vw - 48px)); }
    }
    body.theme-light.notes-visible,
    body.theme-light.previews-visible { padding-top: 0; }
    body.theme-light .web-ui-header { display: none; }
    body.theme-light > .container > .tabs { display: none; }
    body.theme-light .container {
      max-width: var(--tl-band);
      margin-left: auto;
      margin-right: auto;
      background: var(--faire-surface);
      border: 1px solid var(--faire-border);
      border-radius: 0;
      box-shadow: 0 1px 3px rgba(0,0,0,0.06);
      padding: 0;
      /* Space for fixed bottom tab bar (.bottom-tabs is position:fixed) */
      padding-bottom: calc(64px + env(safe-area-inset-bottom, 0px));
      display: flex;
      flex-direction: column;
      width: 100%;
      flex: 1 1 auto;
      min-height: 0;
      height: 100vh;
      max-height: 100dvh;
      overflow: hidden;
      box-sizing: border-box;
    }
    body.theme-light.notes-visible .container,
    body.theme-light.previews-visible .container { max-width: var(--tl-band); }
    /* Must include .active — otherwise this beats .tab-content { display:none } and Remote never hides */
    body.theme-light #tab-remote.tab-content.active {
      display: flex;
      flex-direction: column;
      flex: 1 1 0;
      min-height: 0;
      overflow: hidden;
    }
    body.theme-light #tab-controls.tab-content.active,
    body.theme-light #tab-settings.tab-content.active {
      flex: 1 1 auto;
      min-height: 0;
      overflow-y: auto;
      -webkit-overflow-scrolling: touch;
      /* Match Remote tab horizontal rhythm (.remote-header-compact / .remote-controls use 18px; top 14px) */
      padding: 14px 18px 20px;
      box-sizing: border-box;
      font-family: var(--faire-font-sans);
      color: var(--faire-text);
      -webkit-font-smoothing: antialiased;
    }
    body.theme-light #tab-controls .info,
    body.theme-light #tab-settings .info {
      background: var(--faire-warm);
      border: 1px solid var(--faire-border);
      color: var(--faire-sub);
      border-radius: var(--faire-radius);
      padding: 12px 14px;
      font-size: 13px;
      line-height: 1.5;
      margin-bottom: 16px;
    }
    body.theme-light #tab-controls .controls-section {
      margin-bottom: 0;
      margin-top: 0;
      padding: 16px;
      border-top: none;
      border: 1px solid var(--faire-border);
      border-radius: var(--faire-radius);
      background: var(--faire-surface);
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
    }
    body.theme-light #tab-settings .controls-section {
      margin-bottom: 0;
      margin-top: 0;
      padding: 20px 22px;
      border-top: none;
      border: 1px solid var(--faire-border);
      border-radius: var(--faire-settings-card-radius);
      background: var(--faire-surface);
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
    }
    body.theme-light #tab-controls .controls-section + .controls-section,
    body.theme-light #tab-settings .controls-section + .controls-section {
      margin-top: 14px;
    }
    /* Controls = action strip (compact caps, sans) */
    body.theme-light #tab-controls .controls-section h3 {
      font-family: var(--faire-font-sans);
      font-size: 10.5px;
      letter-spacing: 0.7px;
      text-transform: uppercase;
      font-weight: 600;
      color: var(--faire-sub);
      margin-top: 0;
      margin-bottom: 14px;
    }
    /* Settings = configuration cards (editorial serif title) */
    body.theme-light #tab-settings .controls-section h3 {
      font-family: var(--faire-font-serif);
      font-size: 18px;
      font-weight: 500;
      letter-spacing: normal;
      text-transform: none;
      color: var(--faire-text);
      margin-top: 0;
      margin-bottom: 6px;
      line-height: 1.25;
    }
    body.theme-light #tab-controls .preset-group,
    body.theme-light #tab-settings .preset-group {
      margin-bottom: 16px;
    }
    body.theme-light #tab-controls label,
    body.theme-light #tab-settings label {
      font-weight: 600;
      font-size: 13px;
      color: var(--faire-text);
      margin-bottom: 6px;
    }
    /* Settings: label / control columns on wider viewports (handoff field-row) */
    @media (min-width: 640px) {
      body.theme-light #tab-settings .preset-group:has(> label) {
        display: grid;
        grid-template-columns: minmax(120px, 180px) minmax(0, 1fr);
        gap: 8px 16px;
        align-items: start;
      }
      body.theme-light #tab-settings .preset-group:has(> label) > label {
        grid-column: 1;
        margin-bottom: 0;
        padding-top: 10px;
      }
      body.theme-light #tab-settings .preset-group:has(> label) > *:not(label) {
        grid-column: 2;
        min-width: 0;
      }
      body.theme-light #tab-settings .preset-group:has(> label) > small {
        grid-column: 1 / -1;
        padding-top: 2px;
      }
      body.theme-light #tab-settings .preset-group:has(> label) > #web-backup-ip-list {
        grid-column: 1 / -1;
      }
      body.theme-light #tab-settings .preset-group:has(> label) > button.btn {
        grid-column: 1 / -1;
      }
    }
    body.theme-light #tab-controls input[type="text"],
    body.theme-light #tab-controls input[type="number"],
    body.theme-light #tab-controls input[type="password"],
    body.theme-light #tab-settings input[type="text"],
    body.theme-light #tab-settings input[type="number"],
    body.theme-light #tab-settings input[type="password"] {
      width: 100%;
      font-family: var(--faire-font-sans);
      border: 1px solid var(--faire-border);
      border-radius: var(--faire-radius);
      padding: 10px 12px;
      font-size: 14px;
      color: var(--faire-text);
      background: var(--faire-surface);
      transition: border-color 0.15s ease;
    }
    body.theme-light #tab-controls input[type="text"]:focus,
    body.theme-light #tab-controls input[type="number"]:focus,
    body.theme-light #tab-controls input[type="password"]:focus,
    body.theme-light #tab-settings input[type="text"]:focus,
    body.theme-light #tab-settings input[type="number"]:focus,
    body.theme-light #tab-settings input[type="password"]:focus,
    body.theme-light #tab-settings select:focus {
      outline: none;
      border-color: var(--faire-text);
    }
    body.theme-light #tab-settings select,
    body.theme-light #tab-settings select.input-field {
      width: 100%;
      font-family: var(--faire-font-sans);
      border: 1px solid var(--faire-border);
      border-radius: var(--faire-radius);
      padding: 10px 12px;
      font-size: 14px;
      color: var(--faire-text);
      background: var(--faire-surface);
    }
    body.theme-light #tab-controls small,
    body.theme-light #tab-settings small {
      color: var(--faire-sub) !important;
      font-size: 12px;
      line-height: 1.45;
    }
    body.theme-light #tab-controls .btn,
    body.theme-light #tab-settings .btn {
      width: 100%;
      background: var(--faire-text);
      color: var(--faire-surface);
      border: 1px solid var(--faire-text);
      font-weight: 600;
      font-size: 14px;
      padding: 12px 16px;
      margin-top: 10px;
      box-shadow: none;
      transform: none;
    }
    body.theme-light #tab-controls .btn:hover,
    body.theme-light #tab-settings .btn:hover {
      opacity: 0.88;
    }
    body.theme-light #tab-controls .btn:active,
    body.theme-light #tab-settings .btn:active {
      transform: none;
    }
    body.theme-light #tab-controls .btn-secondary,
    body.theme-light #tab-settings .btn-secondary {
      background: var(--faire-surface);
      color: var(--faire-text);
      border: 1px solid var(--faire-border);
    }
    body.theme-light #tab-controls .btn-secondary:hover,
    body.theme-light #tab-settings .btn-secondary:hover {
      background: var(--faire-warm);
      border-color: var(--faire-border);
      color: var(--faire-text);
      opacity: 1;
    }
    body.theme-light #tab-controls .controls-section:has(#btn-open-presentation) > div:has(#btn-open-presentation) {
      display: flex;
      gap: 10px;
      margin-top: 10px;
    }
    body.theme-light #tab-controls .controls-section:has(#btn-open-presentation) > div:has(#btn-open-presentation) > .btn {
      margin-top: 0;
      flex: 1;
      width: auto;
    }
    body.theme-light #tab-controls .btn-control {
      background: var(--faire-surface);
      border: 1px solid var(--faire-border);
      color: var(--faire-text);
      border-radius: var(--faire-radius);
      font-weight: 500;
      font-size: 13px;
      padding: 12px 14px;
      transform: none;
      box-shadow: none;
    }
    body.theme-light #tab-controls .btn-control:hover {
      background: var(--faire-warm);
      border-color: var(--faire-border);
      color: var(--faire-text);
      transform: none;
      box-shadow: none;
    }
    body.theme-light #tab-controls .controls-grid {
      margin-bottom: 0;
      gap: 10px;
    }
    body.theme-light #tab-settings .web-ui-callout--warning {
      margin-top: 12px;
      padding: 12px 14px;
      background: var(--faire-warm);
      border: 1px solid var(--tmr-warn-bd);
      color: var(--tmr-warn-fg);
      border-radius: var(--faire-radius);
      font-size: 13px;
      line-height: 1.45;
    }
    body.theme-light #tab-settings #web-tunnel-qr-row input[type="number"] {
      width: 72px;
      padding: 8px 10px;
      border: 1px solid var(--faire-border) !important;
      border-radius: var(--faire-radius) !important;
      font-size: 13px;
      font-family: var(--faire-font-sans);
      color: var(--faire-text);
      background: var(--faire-surface);
    }
    body.theme-light #tab-settings #debug-console {
      border: 1px solid var(--faire-border) !important;
      border-radius: var(--faire-radius) !important;
      font-family: var(--faire-font-mono) !important;
    }
    body.theme-light #tab-controls .web-preset-launch-row {
      display: flex;
      gap: 10px;
      align-items: stretch;
      margin-bottom: 10px;
    }
    body.theme-light #tab-controls .web-preset-launch-label {
      font-weight: 600;
      font-size: 10.5px;
      letter-spacing: 0.5px;
      text-transform: uppercase;
      color: var(--faire-sub);
      padding: 12px 0;
      min-width: 120px;
    }
    body.theme-light #tab-controls .web-preset-launch-actions {
      display: flex;
      gap: 10px;
      flex: 1;
      min-width: 0;
    }
    body.theme-light #tab-controls .web-preset-launch-actions .btn {
      margin-top: 0;
      flex: 1;
      width: auto;
    }
    body.theme-light #tab-controls .web-preset-empty {
      color: var(--faire-muted) !important;
      font-style: italic;
      padding: 20px;
      text-align: center;
      font-size: 13px;
      line-height: 1.5;
    }
    body.theme-light #tab-controls .web-preset-empty-link {
      display: inline;
      margin: 0;
      padding: 0;
      border: none;
      background: none;
      font: inherit;
      font-style: normal;
      font-weight: 600;
      color: var(--faire-text);
      text-decoration: underline;
      text-underline-offset: 2px;
      cursor: pointer;
    }
    body.theme-light #tab-controls .web-preset-empty-link:hover {
      opacity: 0.85;
    }
    body.theme-light #tab-settings [data-web-preset-row] {
      display: flex;
      gap: 10px;
      align-items: center;
      margin-bottom: 10px;
      width: 100%;
      min-width: 0;
    }
    body.theme-light #tab-settings [data-web-preset-row] input[data-web-preset-url="true"] {
      flex: 1;
      min-width: 0;
    }
    body.theme-light #tab-settings [data-web-preset-row] .btn {
      margin-top: 0;
      width: auto;
      flex-shrink: 0;
      padding: 10px 14px;
    }
    body.theme-light #tab-settings [data-web-backup-row] {
      display: flex;
      gap: 10px;
      align-items: center;
      width: 100%;
      min-width: 0;
    }
    body.theme-light #tab-settings [data-web-backup-row] .btn {
      margin-top: 0;
      width: auto;
      flex-shrink: 0;
    }
    body.theme-light #tab-settings input[type="checkbox"] + label {
      font-weight: 400 !important;
      margin-bottom: 0 !important;
      color: var(--faire-text);
    }
    body.theme-light .remote-header-compact {
      padding: 14px 18px 10px;
      font-size: 11.5px;
      flex-shrink: 0;
    }
    body.theme-light .remote-machine-name {
      font-size: 14px;
      font-weight: 500;
      color: var(--faire-text);
      flex: 1 1 auto;
      min-width: 0;
    }
    body.theme-light .remote-status-dot { background: #43a047; }
    body.theme-light .slide-counter {
      margin-left: auto;
      color: var(--faire-sub);
      font-size: 11.5px;
      padding-right: 0;
    }
    body.theme-light .notes-toggle-btn,
    body.theme-light .preview-toggle-btn {
      width: 32px;
      height: 32px;
      padding: 0;
      border-radius: var(--faire-radius);
      border: 1px solid var(--faire-border);
      background: var(--faire-surface);
      color: var(--faire-sub);
      justify-content: center;
    }
    body.theme-light .notes-toggle-btn svg,
    body.theme-light .preview-toggle-btn svg {
      width: 16px;
      height: 16px;
    }
    body.theme-light .toggle-btn-text { display: none; }
    body.theme-light .slide-previews-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      padding: 0 18px 12px;
      background: transparent;
      border: none;
    }
    body.theme-light .slide-preview-card {
      display: flex;
      flex-direction: row;
      gap: 8px;
      align-items: center;
      padding: 0;
      background: transparent;
      border: none;
    }
    body.theme-light .slide-preview-img {
      width: 96px;
      height: 54px;
      border-radius: 2px;
      border: 1px solid var(--faire-border);
      object-fit: cover;
      max-width: none;
      max-height: none;
    }
    body.theme-light .slide-preview-card:last-child .slide-preview-img {
      width: 72px;
      height: 40px;
    }
    body.theme-light .slide-preview-label {
      font-size: 10.5px;
      color: var(--faire-sub);
      letter-spacing: 0.7px;
      text-transform: uppercase;
    }
    body.theme-light .speaker-notes-container {
      flex: 1;
      min-height: 0;
      margin: 0 18px;
      background: var(--faire-surface);
      border: 1px solid var(--faire-border);
      border-radius: var(--faire-radius);
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }
    body.theme-light .notes-zoom-controls {
      position: static;
      display: flex;
      visibility: visible;
      padding: 6px 8px 6px 14px;
      background: var(--faire-warm);
      border-bottom: 1px solid var(--faire-border);
      gap: 8px;
      align-items: center;
      margin-top: 0;
      bottom: auto;
      z-index: 1;
    }
    body.theme-light .speaker-notes-toolbar-label {
      font-size: 10.5px;
      letter-spacing: 0.7px;
      text-transform: uppercase;
      color: var(--faire-sub);
      flex: 1;
    }
    body.theme-light .notes-zoom-btn {
      background: var(--faire-surface);
      border: 1px solid var(--faire-border);
      border-radius: var(--faire-radius);
      padding: 0;
      width: 32px;
      height: 30px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      color: var(--faire-text);
      font-size: 16px;
      font-weight: 400;
    }
    body.theme-light .notes-zoom-btn:hover {
      background: var(--faire-warm);
      color: var(--faire-text);
      border-color: var(--faire-border);
    }
    body.theme-light #notes-zoom-readout {
      font-family: var(--faire-font-mono);
      font-size: 11px;
      color: var(--faire-sub);
      padding: 0 6px;
      min-width: 36px;
      text-align: center;
    }
    body.theme-light .speaker-notes-content-wrapper {
      padding: 16px 18px;
      flex: 1;
      overflow-y: auto;
      border-radius: 0;
      border: none;
      background: transparent;
      height: auto;
      min-height: 120px;
    }
    body.theme-light .speaker-notes-content {
      font-family: var(--faire-font-serif);
      font-size: 19px;
      line-height: 30px;
      color: var(--faire-text);
    }
    body.theme-light .remote-controls {
      padding: 14px 18px;
      display: flex;
      gap: 10px;
      flex-shrink: 0;
    }
    body.theme-light .remote-btn {
      flex: 1;
      height: 72px;
      border-radius: var(--faire-radius);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 2px;
      min-height: 72px;
      padding: 12px 16px;
      font-family: var(--faire-font-sans);
    }
    body.theme-light .remote-btn .remote-btn-label {
      margin-top: 2px;
      font-size: 11px;
      letter-spacing: 0.5px;
      text-transform: uppercase;
      font-weight: 600;
    }
    body.theme-light .remote-btn svg {
      width: 28px;
      height: 28px;
    }
    body.theme-light .remote-btn-prev {
      background: var(--faire-surface);
      border: 1px solid var(--faire-border);
      color: var(--faire-text);
    }
    body.theme-light .remote-btn-next {
      background: var(--faire-text);
      border: 1px solid var(--faire-text);
      color: var(--faire-surface);
    }
    body.theme-light .bottom-tabs {
      display: flex;
      position: fixed;
      bottom: 0;
      left: 0;
      right: 0;
      width: var(--tl-band);
      max-width: 100%;
      margin: 0 auto;
      box-sizing: border-box;
      border-top: 1px solid var(--faire-border);
      background: var(--faire-surface);
      padding: 6px 0 calc(12px + env(safe-area-inset-bottom, 0px));
      z-index: 100;
      box-shadow: 0 -2px 12px rgba(0, 0, 0, 0.06);
    }
    body.theme-light .bottom-tabs .tab-btn {
      flex: 1;
      background: transparent;
      border: none;
      border-radius: 0;
      padding: 6px 0;
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 2px;
      color: var(--faire-sub);
      font-size: 10px;
      letter-spacing: 0.2px;
      font-weight: 400;
    }
    body.theme-light .bottom-tabs .tab-btn.active {
      color: var(--faire-text);
      font-weight: 500;
      background: transparent;
    }
    body.theme-light .bottom-tabs .tab-icon {
      width: 22px;
      height: 22px;
    }
    body.theme-light .build-number {
      bottom: calc(70px + env(safe-area-inset-bottom, 0px));
      z-index: 101;
    }
    body.theme-light .tab-btn { border-radius: 0; }
    body.theme-light .btn, body.theme-light .remote-btn, body.theme-light .notes-zoom-btn { border-radius: var(--faire-radius); }
    body.theme-light .slide-previews-grid, body.theme-light .speaker-notes-content-wrapper { border-radius: 0; }
    body.theme-light .stagetimer-container {
      border-radius: var(--faire-radius);
      background: var(--tmr-idle-bg);
      border: 1px solid var(--tmr-idle-bd);
      color: var(--tmr-idle-fg);
      box-shadow: none;
      padding: 12px 14px;
      height: auto;
      text-align: left;
      display: flex;
      flex-direction: column;
      gap: 10px;
      margin: 0 18px 12px;
    }
    body.theme-light .stagetimer-container.running {
      background: var(--tmr-run-bg);
      border-color: var(--tmr-run-bd);
      color: var(--tmr-run-fg);
    }
    body.theme-light .stagetimer-container.warning {
      background: var(--tmr-warn-bg);
      border-color: var(--tmr-warn-bd);
      color: var(--tmr-warn-fg);
    }
    body.theme-light .stagetimer-container.critical {
      background: var(--tmr-crit-bg);
      border-color: var(--tmr-crit-bd);
      color: var(--tmr-crit-fg);
    }
    body.theme-light .stagetimer-container.overtime {
      background: var(--tmr-over-bg);
      border-color: var(--tmr-over-bd);
      color: var(--tmr-over-fg);
    }
    body.theme-light .stagetimer-container.error {
      background: var(--tmr-crit-bg);
      border-color: var(--tmr-crit-bd);
      color: var(--tmr-crit-fg);
    }
    body.theme-light .stagetimer-container.disabled {
      background: var(--tmr-idle-bg);
      border-color: var(--tmr-idle-bd);
      color: var(--faire-muted);
    }
    body.theme-light .stagetimer-time {
      font-family: var(--faire-font-mono);
      font-size: 28px;
      font-weight: 500;
      letter-spacing: 1px;
      font-variant-numeric: tabular-nums;
      color: var(--tmr-idle-clk);
      margin: 0;
    }
    body.theme-light .stagetimer-container.running .stagetimer-time { color: var(--tmr-run-clk); }
    body.theme-light .stagetimer-container.warning .stagetimer-time { color: var(--tmr-warn-clk); }
    body.theme-light .stagetimer-container.critical .stagetimer-time,
    body.theme-light .stagetimer-container.error .stagetimer-time { color: var(--tmr-crit-clk); }
    body.theme-light .stagetimer-container.overtime .stagetimer-time { color: var(--tmr-over-clk); }
    body.theme-light .stagetimer-container.disabled .stagetimer-time { color: var(--faire-muted); }
    body.theme-light .stagetimer-label {
      font-size: 10.5px;
      letter-spacing: 0.7px;
      text-transform: uppercase;
      font-weight: 500;
      opacity: 1;
    }
    body.theme-light .stagetimer-status { font-size: 11.5px; opacity: 1; }
    body.theme-light .stagetimer-messages {
      position: static;
      border-top: 1px solid currentColor;
      background: transparent;
      backdrop-filter: none;
      padding: 10px 0 0;
      max-height: none;
      margin: 0;
    }
    body.theme-light .stagetimer-message {
      background: transparent;
      color: currentColor;
      padding: 0;
      font-size: 12.5px;
      line-height: 17px;
    }
    body.theme-light h1, body.theme-light h2, body.theme-light h3 { color: #212121; }
    @media (max-width: 768px) {
      body.theme-light .container {
        width: min(100%, calc(100vw - 24px));
        max-width: calc(100vw - 24px);
        margin-left: auto;
        margin-right: auto;
      }
      body.theme-light.notes-visible .container,
      body.theme-light.previews-visible .container {
        width: min(100%, calc(100vw - 24px));
        max-width: calc(100vw - 24px);
      }
    }
    /* --- theme-original: semantic tokens + shared language (classic purple brand) --- */
    body.theme-original {
      --ui-section-bg: #ffffff;
      --ui-section-muted-bg: #f4f6fb;
      --ui-border: #e2e6ef;
      --ui-text: #2a2a2a;
      --ui-muted: #6b7280;
      --ui-info-bg: #f0f4ff;
      --ui-info-bd: #c7d2fe;
      --ui-info-fg: #4338ca;
      --ui-accent: #667eea;
      --ui-accent-contrast: #ffffff;
      --ui-secondary-bg: #ffffff;
      --ui-secondary-fg: #374151;
      --ui-secondary-bd: #d1d5db;
      --ui-radius: 4px;
      --ui-radius-lg: 4px;
      --ui-settings-card-radius: 10px;
      --ui-card-padding: 16px;
      --ui-settings-card-padding: 20px 22px;
      --ui-section-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
      --stm-idle-bg: #eef1f7;
      --stm-idle-bd: #d8dde8;
      --stm-idle-fg: #5a6170;
      --stm-idle-clk: #1e2433;
      --stm-run-bg: #e8f5e9;
      --stm-run-bd: #a5d6a7;
      --stm-run-fg: #2e7d32;
      --stm-run-clk: #1b5e20;
      --stm-warn-bg: #fff8e1;
      --stm-warn-bd: #ffe082;
      --stm-warn-fg: #e65100;
      --stm-warn-clk: #bf360c;
      --stm-crit-bg: #ffebee;
      --stm-crit-bd: #ef9a9a;
      --stm-crit-fg: #c62828;
      --stm-crit-clk: #b71c1c;
      --stm-over-bg: #3e2723;
      --stm-over-bd: #6d4c41;
      --stm-over-fg: #ffccbc;
      --stm-over-clk: #ffffff;
      --stm-dis-bg: #eceff1;
      --stm-dis-bd: #cfd8dc;
      --stm-dis-fg: #78909c;
      --stm-dis-clk: #90a4ae;
    }
    body.theme-original .container { border-radius: var(--faire-radius); }
    /* Dark: glass + cool tokens */
    body.theme-dark {
      --ui-section-bg: rgba(0, 0, 0, 0.32);
      --ui-section-muted-bg: rgba(255, 255, 255, 0.04);
      --ui-border: rgba(255, 255, 255, 0.14);
      --ui-text: rgba(255, 255, 255, 0.92);
      --ui-muted: rgba(255, 255, 255, 0.62);
      --ui-info-bg: rgba(99, 102, 241, 0.15);
      --ui-info-bd: rgba(129, 140, 248, 0.35);
      --ui-info-fg: rgba(199, 210, 254, 0.95);
      --ui-accent: #6366f1;
      --ui-accent-contrast: #ffffff;
      --ui-secondary-bg: rgba(255, 255, 255, 0.08);
      --ui-secondary-fg: rgba(255, 255, 255, 0.9);
      --ui-secondary-bd: rgba(255, 255, 255, 0.2);
      --ui-radius: 4px;
      --ui-radius-lg: 4px;
      --ui-settings-card-radius: 10px;
      --ui-card-padding: 16px;
      --ui-settings-card-padding: 18px 20px;
      --ui-section-shadow: none;
      --stm-idle-bg: rgba(28, 28, 30, 0.75);
      --stm-idle-bd: rgba(255, 255, 255, 0.12);
      --stm-idle-fg: rgba(255, 255, 255, 0.72);
      --stm-idle-clk: rgba(255, 255, 255, 0.95);
      --stm-run-bg: rgba(46, 125, 50, 0.38);
      --stm-run-bd: rgba(129, 199, 132, 0.45);
      --stm-run-fg: #e8f5e9;
      --stm-run-clk: #ffffff;
      --stm-warn-bg: rgba(230, 126, 34, 0.28);
      --stm-warn-bd: rgba(255, 183, 77, 0.45);
      --stm-warn-fg: #ffe0b2;
      --stm-warn-clk: #fff8e1;
      --stm-crit-bg: rgba(198, 40, 40, 0.38);
      --stm-crit-bd: rgba(239, 83, 80, 0.55);
      --stm-crit-fg: #ffcdd2;
      --stm-crit-clk: #ffffff;
      --stm-over-bg: rgba(62, 39, 35, 0.92);
      --stm-over-bd: rgba(141, 110, 99, 0.7);
      --stm-over-fg: #ffccbc;
      --stm-over-clk: #ffffff;
      --stm-dis-bg: rgba(55, 55, 58, 0.6);
      --stm-dis-bd: rgba(255, 255, 255, 0.1);
      --stm-dis-fg: rgba(255, 255, 255, 0.45);
      --stm-dis-clk: rgba(255, 255, 255, 0.5);
      background: linear-gradient(180deg, #1c1c1e 0%, #2c2c2e 100%);
      padding-top: 25vh;
    }
    body.theme-dark.notes-visible,
    body.theme-dark.previews-visible { padding-top: 4%; }
    body.theme-dark .container { background: rgba(44, 44, 46, 0.72); backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px); border-radius: var(--faire-radius); border: 1px solid rgba(255,255,255,0.08); box-shadow: 0 8px 32px rgba(0,0,0,0.4); max-width: 75%; }
    body.theme-dark.notes-visible .container,
    body.theme-dark.previews-visible .container { max-width: 85%; }
    body.theme-dark h1, body.theme-dark h2, body.theme-dark h3 { color: rgba(255,255,255,0.92); }
    body.theme-dark .tab-btn { background: rgba(255,255,255,0.1); color: rgba(255,255,255,0.9); border: 1px solid rgba(255,255,255,0.15); }
    body.theme-dark .tab-btn.active { background: rgba(255,255,255,0.2); }
    body.theme-dark .remote-btn-prev, body.theme-dark .remote-btn-next { background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.2); }
    body.theme-dark .remote-btn:hover { background: rgba(255,255,255,0.25); }
    body.theme-dark .slide-previews-grid, body.theme-dark .speaker-notes-content-wrapper { background: rgba(0,0,0,0.2); border: 1px solid rgba(255,255,255,0.1); border-radius: var(--faire-radius); }
    body.theme-dark .speaker-notes-content { color: rgba(255,255,255,0.88); }
    body.theme-dark .notes-zoom-btn { background: rgba(255,255,255,0.1); color: rgba(255,255,255,0.9); border-color: rgba(255,255,255,0.15); }
    body.theme-dark .remote-header-compact .remote-machine-name { color: rgba(255,255,255,0.9); }
    body.theme-dark .remote-header-compact .slide-counter { color: rgba(255,255,255,0.55); }
    @media (max-width: 768px) {
      body.theme-dark .container { width: min(100%, calc(100vw - 24px)); max-width: calc(100vw - 24px); margin-left: auto; margin-right: auto; }
      body.theme-dark.notes-visible .container, body.theme-dark.previews-visible .container { width: min(100%, calc(100vw - 24px)); max-width: calc(100vw - 24px); }
    }
    /* Max: dense full-viewport */
    body.theme-max {
      --ui-section-bg: rgba(255, 255, 255, 0.92);
      --ui-section-muted-bg: #f1f5f9;
      --ui-border: #dde1e7;
      --ui-text: #1a1a1a;
      --ui-muted: #64748b;
      --ui-info-bg: #eff6ff;
      --ui-info-bd: #bfdbfe;
      --ui-info-fg: #1d4ed8;
      --ui-accent: #4f46e5;
      --ui-accent-contrast: #ffffff;
      --ui-secondary-bg: #f8fafc;
      --ui-secondary-fg: #334155;
      --ui-secondary-bd: #cbd5e1;
      --ui-radius: 4px;
      --ui-radius-lg: 4px;
      --ui-settings-card-radius: 10px;
      --ui-card-padding: 10px 12px;
      --ui-settings-card-padding: 12px 14px;
      --ui-section-shadow: none;
      --stm-idle-bg: #f1f5f9;
      --stm-idle-bd: #cbd5e1;
      --stm-idle-fg: #475569;
      --stm-idle-clk: #0f172a;
      --stm-run-bg: #dcfce7;
      --stm-run-bd: #86efac;
      --stm-run-fg: #166534;
      --stm-run-clk: #14532d;
      --stm-warn-bg: #fef9c3;
      --stm-warn-bd: #fde047;
      --stm-warn-fg: #a16207;
      --stm-warn-clk: #713f12;
      --stm-crit-bg: #fee2e2;
      --stm-crit-bd: #fca5a5;
      --stm-crit-fg: #b91c1c;
      --stm-crit-clk: #7f1d1d;
      --stm-over-bg: #292524;
      --stm-over-bd: #57534e;
      --stm-over-fg: #fecaca;
      --stm-over-clk: #ffffff;
      --stm-dis-bg: #e2e8f0;
      --stm-dis-bd: #cbd5e1;
      --stm-dis-fg: #94a3b8;
      --stm-dis-clk: #64748b;
    }
    body.theme-max .container { max-width: 100%; width: 100%; height: 100vh; max-height: 100vh; padding: 8px 12px; border-radius: 0; display: flex; flex-direction: column; }
    body.theme-max h1 { font-size: 14px; padding: 4px 0; }
    body.theme-max .system-icon { width: 18px; height: 18px; }
    body.theme-max .tabs { padding: 4px 0; }
    body.theme-max .tab-btn { padding: 6px 14px; font-size: 13px; }
    /* Must include .active — otherwise display:flex beats .tab-content { display:none } and Remote leaks onto other tabs */
    body.theme-max #tab-remote.tab-content.active { display: flex; flex-direction: column; flex: 1; min-height: 0; }
    body.theme-max .remote-header { flex-shrink: 0; padding: 4px 0; }
    body.theme-max .remote-header h2 { font-size: 14px; }
    body.theme-max .slide-previews-container { order: 1; flex: 1 1 auto; min-height: 0; margin-top: 4px; }
    body.theme-max .speaker-notes-container { order: 2; flex: 1 1 auto; min-height: 0; margin-top: 4px; }
    body.theme-max .speaker-notes-content-wrapper { height: 120px; min-height: 80px; }
    body.theme-max .remote-controls { order: 3; flex-shrink: 0; margin-top: 8px; }
    body.theme-max .remote-btn { min-height: 52px; padding: 12px 16px; font-size: 16px; }
    /* Touch: soft large-radius */
    body.theme-touch {
      --ui-section-bg: #ffffff;
      --ui-section-muted-bg: #f8fafc;
      --ui-border: #e2e8f0;
      --ui-text: #1e293b;
      --ui-muted: #64748b;
      --ui-info-bg: #f0fdf4;
      --ui-info-bd: #bbf7d0;
      --ui-info-fg: #166534;
      --ui-accent: #4f46e5;
      --ui-accent-contrast: #ffffff;
      --ui-secondary-bg: #f8fafc;
      --ui-secondary-fg: #334155;
      --ui-secondary-bd: #e2e8f0;
      --ui-radius: 4px;
      --ui-radius-lg: 4px;
      --ui-settings-card-radius: 10px;
      --ui-card-padding: 18px;
      --ui-settings-card-padding: 22px 24px;
      --ui-section-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);
      --stm-idle-bg: #f8fafc;
      --stm-idle-bd: #e2e8f0;
      --stm-idle-fg: #64748b;
      --stm-idle-clk: #0f172a;
      --stm-run-bg: #ecfdf5;
      --stm-run-bd: #6ee7b7;
      --stm-run-fg: #047857;
      --stm-run-clk: #065f46;
      --stm-warn-bg: #fffbeb;
      --stm-warn-bd: #fcd34d;
      --stm-warn-fg: #b45309;
      --stm-warn-clk: #92400e;
      --stm-crit-bg: #fef2f2;
      --stm-crit-bd: #fca5a5;
      --stm-crit-fg: #b91c1c;
      --stm-crit-clk: #991b1b;
      --stm-over-bg: #431407;
      --stm-over-bd: #9a3412;
      --stm-over-fg: #ffedd5;
      --stm-over-clk: #ffffff;
      --stm-dis-bg: #f1f5f9;
      --stm-dis-bd: #e2e8f0;
      --stm-dis-fg: #94a3b8;
      --stm-dis-clk: #94a3b8;
      background: #e8eaf0;
      padding: 16px;
    }
    body.theme-touch .container { background: #fff; border-radius: var(--faire-radius); box-shadow: 0 8px 32px rgba(0,0,0,0.12); max-width: 90%; padding: 28px; }
    body.theme-touch .remote-btn { min-height: 80px; padding: 24px 28px; font-size: 22px; -webkit-tap-highlight-color: transparent; }
    body.theme-touch .remote-btn:active { transform: scale(0.97); }
    body.theme-touch .tab-btn { padding: 16px 28px; font-size: 18px; min-height: 52px; -webkit-tap-highlight-color: transparent; }
    body.theme-touch .notes-toggle-btn, body.theme-touch .preview-toggle-btn { padding: 14px 20px; font-size: 16px; min-height: 48px; }
    body.theme-touch .slide-previews-grid, body.theme-touch .speaker-notes-content-wrapper { border-radius: var(--faire-radius); padding: 18px; }
    @media (max-width: 768px) {
      body.theme-touch .container { width: min(100%, calc(100vw - 24px)); max-width: calc(100vw - 24px); margin-left: auto; margin-right: auto; }
    }
    /* Thumb: slate stage */
    body.theme-thumb {
      --ui-section-bg: rgba(0, 0, 0, 0.22);
      --ui-section-muted-bg: rgba(255, 255, 255, 0.06);
      --ui-border: rgba(255, 255, 255, 0.12);
      --ui-text: rgba(255, 255, 255, 0.92);
      --ui-muted: rgba(255, 255, 255, 0.6);
      --ui-info-bg: rgba(56, 189, 248, 0.12);
      --ui-info-bd: rgba(125, 211, 252, 0.3);
      --ui-info-fg: #bae6fd;
      --ui-accent: #7dd3fc;
      --ui-accent-contrast: #0c4a6e;
      --ui-secondary-bg: rgba(255, 255, 255, 0.08);
      --ui-secondary-fg: rgba(255, 255, 255, 0.9);
      --ui-secondary-bd: rgba(255, 255, 255, 0.18);
      --ui-radius: 4px;
      --ui-radius-lg: 4px;
      --ui-settings-card-radius: 10px;
      --ui-card-padding: 16px;
      --ui-settings-card-padding: 18px 20px;
      --ui-section-shadow: none;
      --stm-idle-bg: rgba(15, 23, 42, 0.55);
      --stm-idle-bd: rgba(148, 163, 184, 0.25);
      --stm-idle-fg: rgba(226, 232, 240, 0.85);
      --stm-idle-clk: #f8fafc;
      --stm-run-bg: rgba(22, 101, 52, 0.45);
      --stm-run-bd: rgba(74, 222, 128, 0.35);
      --stm-run-fg: #dcfce7;
      --stm-run-clk: #ffffff;
      --stm-warn-bg: rgba(180, 83, 9, 0.35);
      --stm-warn-bd: rgba(251, 191, 36, 0.4);
      --stm-warn-fg: #fef3c7;
      --stm-warn-clk: #fffbeb;
      --stm-crit-bg: rgba(153, 27, 27, 0.45);
      --stm-crit-bd: rgba(248, 113, 113, 0.45);
      --stm-crit-fg: #fecaca;
      --stm-crit-clk: #ffffff;
      --stm-over-bg: rgba(30, 20, 18, 0.95);
      --stm-over-bd: rgba(248, 113, 113, 0.35);
      --stm-over-fg: #ffedd5;
      --stm-over-clk: #ffffff;
      --stm-dis-bg: rgba(30, 41, 59, 0.5);
      --stm-dis-bd: rgba(148, 163, 184, 0.2);
      --stm-dis-fg: rgba(148, 163, 184, 0.75);
      --stm-dis-clk: rgba(203, 213, 225, 0.8);
      background: linear-gradient(180deg, #2d3748 0%, #1a202c 100%);
      padding: 12px;
      min-height: 100vh;
    }
    body.theme-thumb .container { background: rgba(255,255,255,0.06); border-radius: var(--faire-radius); max-width: 96%; padding: 16px; display: flex; flex-direction: column; min-height: 90vh; }
    body.theme-thumb h1 { font-size: 16px; color: rgba(255,255,255,0.85); padding: 6px 0; }
    body.theme-thumb .tabs { margin-bottom: 12px; }
    body.theme-thumb .tab-btn { padding: 12px 20px; font-size: 15px; color: rgba(255,255,255,0.9); background: rgba(255,255,255,0.08); -webkit-tap-highlight-color: transparent; }
    body.theme-thumb .tab-btn.active { background: rgba(255,255,255,0.2); }
    body.theme-thumb #tab-remote.tab-content.active { display: flex; flex-direction: column; flex: 1; min-height: 0; }
    body.theme-thumb .slide-previews-container { order: 1; flex: 1 1 auto; min-height: 0; }
    body.theme-thumb .speaker-notes-container { order: 2; flex: 1 1 auto; min-height: 0; }
    body.theme-thumb .remote-controls { order: 3; flex-shrink: 0; margin-top: 12px; }
    body.theme-thumb .remote-btn { min-height: 72px; padding: 20px 24px; font-size: 20px; -webkit-tap-highlight-color: transparent; }
    body.theme-thumb .remote-btn-prev, body.theme-thumb .remote-btn-next { background: rgba(255,255,255,0.2); color: #fff; border: 1px solid rgba(255,255,255,0.3); }
    body.theme-thumb .remote-btn:hover { background: rgba(255,255,255,0.35); }
    body.theme-thumb .remote-btn:active { transform: scale(0.98); }
    body.theme-thumb .slide-previews-grid, body.theme-thumb .speaker-notes-content-wrapper { background: rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.1); border-radius: var(--faire-radius); color: rgba(255,255,255,0.9); }
    body.theme-thumb .speaker-notes-content { color: rgba(255,255,255,0.9); }
    body.theme-thumb .notes-toggle-btn, body.theme-thumb .preview-toggle-btn { background: rgba(255,255,255,0.12); color: #fff; border: 1px solid rgba(255,255,255,0.2); }
    body.theme-thumb .remote-header-compact .remote-machine-name { color: rgba(255,255,255,0.9); }
    body.theme-thumb .remote-header-compact .slide-counter { color: rgba(255,255,255,0.55); }
    @media (max-width: 768px) {
      body.theme-thumb .container { width: min(100%, calc(100vw - 24px)); max-width: calc(100vw - 24px); margin-left: auto; margin-right: auto; }
    }

    /* Shared: align Remote / notes / previews corner radii with light (--faire-radius / preview imgs 2px) */
    body.theme-original .remote-btn,
    body.theme-dark .remote-btn,
    body.theme-max .remote-btn,
    body.theme-touch .remote-btn,
    body.theme-thumb .remote-btn,
    body.theme-original .notes-toggle-btn,
    body.theme-dark .notes-toggle-btn,
    body.theme-max .notes-toggle-btn,
    body.theme-touch .notes-toggle-btn,
    body.theme-thumb .notes-toggle-btn,
    body.theme-original .preview-toggle-btn,
    body.theme-dark .preview-toggle-btn,
    body.theme-max .preview-toggle-btn,
    body.theme-touch .preview-toggle-btn,
    body.theme-thumb .preview-toggle-btn,
    body.theme-original .notes-zoom-btn,
    body.theme-dark .notes-zoom-btn,
    body.theme-max .notes-zoom-btn,
    body.theme-touch .notes-zoom-btn,
    body.theme-thumb .notes-zoom-btn,
    body.theme-original .tab-btn,
    body.theme-dark .tab-btn,
    body.theme-max .tab-btn,
    body.theme-touch .tab-btn,
    body.theme-thumb .tab-btn {
      border-radius: var(--faire-radius);
    }
    body.theme-original .slide-preview-card.clickable,
    body.theme-dark .slide-preview-card.clickable,
    body.theme-max .slide-preview-card.clickable,
    body.theme-touch .slide-preview-card.clickable,
    body.theme-thumb .slide-preview-card.clickable {
      border-radius: var(--faire-radius);
    }
    body.theme-original .slide-preview-img,
    body.theme-dark .slide-preview-img,
    body.theme-max .slide-preview-img,
    body.theme-touch .slide-preview-img,
    body.theme-thumb .slide-preview-img {
      border-radius: 2px;
    }
    body.theme-original .slide-previews-grid,
    body.theme-original .speaker-notes-content-wrapper,
    body.theme-max .slide-previews-grid,
    body.theme-max .speaker-notes-content-wrapper {
      border-radius: var(--faire-radius);
    }
    body.theme-original .speaker-notes-container,
    body.theme-dark .speaker-notes-container,
    body.theme-max .speaker-notes-container,
    body.theme-touch .speaker-notes-container,
    body.theme-thumb .speaker-notes-container {
      border-radius: var(--faire-radius);
    }

    /* Shared: flat stagetimer (non-light themes) */
    body.theme-original .stagetimer-container,
    body.theme-dark .stagetimer-container,
    body.theme-max .stagetimer-container,
    body.theme-touch .stagetimer-container,
    body.theme-thumb .stagetimer-container {
      height: auto;
      min-height: 0;
      overflow: visible;
      position: relative;
      text-align: left;
      display: flex;
      flex-direction: column;
      gap: 10px;
      padding: 12px 14px;
      margin-bottom: 16px;
      border-radius: var(--ui-radius);
      box-shadow: none;
      background: var(--stm-idle-bg);
      border: 1px solid var(--stm-idle-bd);
      color: var(--stm-idle-fg);
    }
    body.theme-original .stagetimer-container.running,
    body.theme-dark .stagetimer-container.running,
    body.theme-max .stagetimer-container.running,
    body.theme-touch .stagetimer-container.running,
    body.theme-thumb .stagetimer-container.running {
      background: var(--stm-run-bg);
      border-color: var(--stm-run-bd);
      color: var(--stm-run-fg);
    }
    body.theme-original .stagetimer-container.warning,
    body.theme-dark .stagetimer-container.warning,
    body.theme-max .stagetimer-container.warning,
    body.theme-touch .stagetimer-container.warning,
    body.theme-thumb .stagetimer-container.warning {
      background: var(--stm-warn-bg);
      border-color: var(--stm-warn-bd);
      color: var(--stm-warn-fg);
    }
    body.theme-original .stagetimer-container.critical,
    body.theme-dark .stagetimer-container.critical,
    body.theme-max .stagetimer-container.critical,
    body.theme-touch .stagetimer-container.critical,
    body.theme-thumb .stagetimer-container.critical {
      background: var(--stm-crit-bg);
      border-color: var(--stm-crit-bd);
      color: var(--stm-crit-fg);
    }
    body.theme-original .stagetimer-container.overtime,
    body.theme-dark .stagetimer-container.overtime,
    body.theme-max .stagetimer-container.overtime,
    body.theme-touch .stagetimer-container.overtime,
    body.theme-thumb .stagetimer-container.overtime {
      background: var(--stm-over-bg);
      border-color: var(--stm-over-bd);
      color: var(--stm-over-fg);
    }
    body.theme-original .stagetimer-container.error,
    body.theme-dark .stagetimer-container.error,
    body.theme-max .stagetimer-container.error,
    body.theme-touch .stagetimer-container.error,
    body.theme-thumb .stagetimer-container.error {
      background: var(--stm-crit-bg);
      border-color: var(--stm-crit-bd);
      color: var(--stm-crit-fg);
    }
    body.theme-original .stagetimer-container.disabled,
    body.theme-dark .stagetimer-container.disabled,
    body.theme-max .stagetimer-container.disabled,
    body.theme-touch .stagetimer-container.disabled,
    body.theme-thumb .stagetimer-container.disabled {
      background: var(--stm-dis-bg);
      border-color: var(--stm-dis-bd);
      color: var(--stm-dis-fg);
    }
    body.theme-original .stagetimer-time,
    body.theme-dark .stagetimer-time,
    body.theme-max .stagetimer-time,
    body.theme-touch .stagetimer-time,
    body.theme-thumb .stagetimer-time {
      font-family: var(--faire-font-mono);
      font-size: 28px;
      font-weight: 600;
      letter-spacing: 1px;
      font-variant-numeric: tabular-nums;
      margin: 0;
      line-height: 1.15;
      color: var(--stm-idle-clk);
    }
    body.theme-original .stagetimer-container.running .stagetimer-time,
    body.theme-dark .stagetimer-container.running .stagetimer-time,
    body.theme-max .stagetimer-container.running .stagetimer-time,
    body.theme-touch .stagetimer-container.running .stagetimer-time,
    body.theme-thumb .stagetimer-container.running .stagetimer-time { color: var(--stm-run-clk); }
    body.theme-original .stagetimer-container.warning .stagetimer-time,
    body.theme-dark .stagetimer-container.warning .stagetimer-time,
    body.theme-max .stagetimer-container.warning .stagetimer-time,
    body.theme-touch .stagetimer-container.warning .stagetimer-time,
    body.theme-thumb .stagetimer-container.warning .stagetimer-time { color: var(--stm-warn-clk); }
    body.theme-original .stagetimer-container.critical .stagetimer-time,
    body.theme-dark .stagetimer-container.critical .stagetimer-time,
    body.theme-max .stagetimer-container.critical .stagetimer-time,
    body.theme-touch .stagetimer-container.critical .stagetimer-time,
    body.theme-thumb .stagetimer-container.critical .stagetimer-time,
    body.theme-original .stagetimer-container.error .stagetimer-time,
    body.theme-dark .stagetimer-container.error .stagetimer-time,
    body.theme-max .stagetimer-container.error .stagetimer-time,
    body.theme-touch .stagetimer-container.error .stagetimer-time,
    body.theme-thumb .stagetimer-container.error .stagetimer-time { color: var(--stm-crit-clk); }
    body.theme-original .stagetimer-container.overtime .stagetimer-time,
    body.theme-dark .stagetimer-container.overtime .stagetimer-time,
    body.theme-max .stagetimer-container.overtime .stagetimer-time,
    body.theme-touch .stagetimer-container.overtime .stagetimer-time,
    body.theme-thumb .stagetimer-container.overtime .stagetimer-time { color: var(--stm-over-clk); }
    body.theme-original .stagetimer-container.disabled .stagetimer-time,
    body.theme-dark .stagetimer-container.disabled .stagetimer-time,
    body.theme-max .stagetimer-container.disabled .stagetimer-time,
    body.theme-touch .stagetimer-container.disabled .stagetimer-time,
    body.theme-thumb .stagetimer-container.disabled .stagetimer-time { color: var(--stm-dis-clk); }
    body.theme-max .stagetimer-time { font-size: 22px; }
    body.theme-original .stagetimer-label,
    body.theme-dark .stagetimer-label,
    body.theme-max .stagetimer-label,
    body.theme-touch .stagetimer-label,
    body.theme-thumb .stagetimer-label {
      font-size: 10.5px;
      letter-spacing: 0.55px;
      text-transform: uppercase;
      font-weight: 600;
      opacity: 1;
      margin-bottom: 0;
    }
    body.theme-original .stagetimer-status,
    body.theme-dark .stagetimer-status,
    body.theme-max .stagetimer-status,
    body.theme-touch .stagetimer-status,
    body.theme-thumb .stagetimer-status { font-size: 12px; margin-top: 0; opacity: 1; }
    body.theme-original .stagetimer-messages,
    body.theme-dark .stagetimer-messages,
    body.theme-max .stagetimer-messages,
    body.theme-touch .stagetimer-messages,
    body.theme-thumb .stagetimer-messages {
      position: static;
      border-top: 1px solid currentColor;
      background: transparent;
      backdrop-filter: none;
      padding: 10px 0 0;
      max-height: none;
      margin: 0;
      border-radius: 0;
    }
    body.theme-original .stagetimer-message,
    body.theme-dark .stagetimer-message,
    body.theme-max .stagetimer-message,
    body.theme-touch .stagetimer-message,
    body.theme-thumb .stagetimer-message {
      background: transparent;
      color: currentColor;
      padding: 0;
      font-size: 12.5px;
      line-height: 1.45;
      margin-bottom: 6px;
      backdrop-filter: none;
    }

    /* Shared: Controls / Settings cards + buttons (non-light) */
    body.theme-original #tab-controls .controls-section,
    body.theme-original #tab-settings .controls-section,
    body.theme-dark #tab-controls .controls-section,
    body.theme-dark #tab-settings .controls-section,
    body.theme-max #tab-controls .controls-section,
    body.theme-max #tab-settings .controls-section,
    body.theme-touch #tab-controls .controls-section,
    body.theme-touch #tab-settings .controls-section,
    body.theme-thumb #tab-controls .controls-section,
    body.theme-thumb #tab-settings .controls-section {
      margin-top: 0;
      margin-bottom: 0;
      padding: var(--ui-card-padding);
      border: 1px solid var(--ui-border);
      border-radius: var(--ui-radius);
      background: var(--ui-section-bg);
      box-shadow: var(--ui-section-shadow);
    }
    body.theme-original #tab-settings .controls-section,
    body.theme-dark #tab-settings .controls-section,
    body.theme-touch #tab-settings .controls-section,
    body.theme-thumb #tab-settings .controls-section {
      padding: var(--ui-settings-card-padding);
      border-radius: var(--ui-settings-card-radius);
    }
    body.theme-max #tab-settings .controls-section {
      padding: var(--ui-settings-card-padding);
      border-radius: var(--ui-settings-card-radius);
    }
    body.theme-original #tab-controls .controls-section + .controls-section,
    body.theme-original #tab-settings .controls-section + .controls-section,
    body.theme-dark #tab-controls .controls-section + .controls-section,
    body.theme-dark #tab-settings .controls-section + .controls-section,
    body.theme-max #tab-controls .controls-section + .controls-section,
    body.theme-max #tab-settings .controls-section + .controls-section,
    body.theme-touch #tab-controls .controls-section + .controls-section,
    body.theme-touch #tab-settings .controls-section + .controls-section,
    body.theme-thumb #tab-controls .controls-section + .controls-section,
    body.theme-thumb #tab-settings .controls-section + .controls-section {
      margin-top: 14px;
    }
    body.theme-original #tab-controls .info,
    body.theme-original #tab-settings .info,
    body.theme-dark #tab-controls .info,
    body.theme-dark #tab-settings .info,
    body.theme-max #tab-controls .info,
    body.theme-max #tab-settings .info,
    body.theme-touch #tab-controls .info,
    body.theme-touch #tab-settings .info,
    body.theme-thumb #tab-controls .info,
    body.theme-thumb #tab-settings .info {
      background: var(--ui-info-bg);
      border: 1px solid var(--ui-info-bd);
      color: var(--ui-info-fg);
      border-radius: var(--ui-radius);
    }
    body.theme-original #tab-controls .controls-section h3,
    body.theme-original #tab-settings .controls-section h3,
    body.theme-dark #tab-controls .controls-section h3,
    body.theme-dark #tab-settings .controls-section h3,
    body.theme-touch #tab-controls .controls-section h3,
    body.theme-touch #tab-settings .controls-section h3,
    body.theme-thumb #tab-controls .controls-section h3,
    body.theme-thumb #tab-settings .controls-section h3,
    body.theme-max #tab-controls .controls-section h3 {
      font-family: var(--faire-font-sans);
      font-size: 10.5px;
      letter-spacing: 0.65px;
      text-transform: uppercase;
      font-weight: 600;
      color: var(--ui-muted);
      margin-top: 0;
      margin-bottom: 12px;
    }
    body.theme-original #tab-settings .controls-section h3,
    body.theme-dark #tab-settings .controls-section h3,
    body.theme-touch #tab-settings .controls-section h3,
    body.theme-thumb #tab-settings .controls-section h3 {
      font-family: var(--faire-font-serif);
      font-size: 18px;
      font-weight: 500;
      letter-spacing: normal;
      text-transform: none;
      color: var(--ui-text);
      margin-bottom: 6px;
      line-height: 1.25;
    }
    body.theme-max #tab-settings .controls-section h3 {
      font-family: var(--faire-font-sans);
      font-size: 11px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: var(--ui-muted);
      margin-bottom: 8px;
    }
    body.theme-original #tab-controls label,
    body.theme-original #tab-settings label,
    body.theme-dark #tab-controls label,
    body.theme-dark #tab-settings label,
    body.theme-max #tab-controls label,
    body.theme-max #tab-settings label,
    body.theme-touch #tab-controls label,
    body.theme-touch #tab-settings label,
    body.theme-thumb #tab-controls label,
    body.theme-thumb #tab-settings label {
      color: var(--ui-text);
    }
    body.theme-original #tab-controls small,
    body.theme-original #tab-settings small,
    body.theme-dark #tab-controls small,
    body.theme-dark #tab-settings small,
    body.theme-max #tab-controls small,
    body.theme-max #tab-settings small,
    body.theme-touch #tab-controls small,
    body.theme-touch #tab-settings small,
    body.theme-thumb #tab-controls small,
    body.theme-thumb #tab-settings small {
      color: var(--ui-muted) !important;
    }
    body.theme-original #tab-controls .btn,
    body.theme-original #tab-settings .btn,
    body.theme-dark #tab-controls .btn,
    body.theme-dark #tab-settings .btn,
    body.theme-max #tab-controls .btn,
    body.theme-max #tab-settings .btn,
    body.theme-touch #tab-controls .btn,
    body.theme-touch #tab-settings .btn,
    body.theme-thumb #tab-controls .btn,
    body.theme-thumb #tab-settings .btn {
      background: var(--ui-accent);
      color: var(--ui-accent-contrast);
      border: 1px solid var(--ui-accent);
      border-radius: var(--ui-radius);
      box-shadow: none;
    }
    body.theme-original #tab-controls .btn:hover,
    body.theme-original #tab-settings .btn:hover,
    body.theme-dark #tab-controls .btn:hover,
    body.theme-dark #tab-settings .btn:hover,
    body.theme-max #tab-controls .btn:hover,
    body.theme-max #tab-settings .btn:hover,
    body.theme-touch #tab-controls .btn:hover,
    body.theme-touch #tab-settings .btn:hover,
    body.theme-thumb #tab-controls .btn:hover,
    body.theme-thumb #tab-settings .btn:hover {
      opacity: 0.92;
      filter: brightness(0.98);
    }
    body.theme-original #tab-controls .btn-secondary,
    body.theme-original #tab-settings .btn-secondary,
    body.theme-dark #tab-controls .btn-secondary,
    body.theme-dark #tab-settings .btn-secondary,
    body.theme-max #tab-controls .btn-secondary,
    body.theme-max #tab-settings .btn-secondary,
    body.theme-touch #tab-controls .btn-secondary,
    body.theme-touch #tab-settings .btn-secondary,
    body.theme-thumb #tab-controls .btn-secondary,
    body.theme-thumb #tab-settings .btn-secondary {
      background: var(--ui-secondary-bg);
      color: var(--ui-secondary-fg);
      border: 1px solid var(--ui-secondary-bd);
    }
    body.theme-original #tab-controls .btn-control,
    body.theme-dark #tab-controls .btn-control,
    body.theme-max #tab-controls .btn-control,
    body.theme-touch #tab-controls .btn-control,
    body.theme-thumb #tab-controls .btn-control {
      background: var(--ui-secondary-bg);
      color: var(--ui-text);
      border: 1px solid var(--ui-border);
      border-radius: var(--ui-radius);
      box-shadow: none;
      transform: none;
    }
    body.theme-original #tab-controls .btn-control:hover,
    body.theme-dark #tab-controls .btn-control:hover,
    body.theme-max #tab-controls .btn-control:hover,
    body.theme-touch #tab-controls .btn-control:hover,
    body.theme-thumb #tab-controls .btn-control:hover {
      background: var(--ui-section-muted-bg, var(--ui-info-bg));
      transform: none;
      box-shadow: none;
    }
    body.theme-dark #tab-controls .btn-control:hover,
    body.theme-thumb #tab-controls .btn-control:hover {
      background: rgba(255, 255, 255, 0.1);
    }
    body.theme-original #tab-settings .web-ui-callout--warning,
    body.theme-dark #tab-settings .web-ui-callout--warning,
    body.theme-max #tab-settings .web-ui-callout--warning,
    body.theme-touch #tab-settings .web-ui-callout--warning,
    body.theme-thumb #tab-settings .web-ui-callout--warning {
      background: var(--stm-warn-bg);
      border: 1px solid var(--stm-warn-bd);
      color: var(--stm-warn-fg);
      border-radius: var(--ui-radius);
    }
    body.theme-original #tab-controls input[type="text"],
    body.theme-original #tab-controls input[type="number"],
    body.theme-original #tab-controls input[type="password"],
    body.theme-original #tab-settings input[type="text"],
    body.theme-original #tab-settings input[type="number"],
    body.theme-original #tab-settings input[type="password"],
    body.theme-dark #tab-controls input[type="text"],
    body.theme-dark #tab-controls input[type="number"],
    body.theme-dark #tab-controls input[type="password"],
    body.theme-dark #tab-settings input[type="text"],
    body.theme-dark #tab-settings input[type="number"],
    body.theme-dark #tab-settings input[type="password"],
    body.theme-max #tab-controls input[type="text"],
    body.theme-max #tab-controls input[type="number"],
    body.theme-max #tab-controls input[type="password"],
    body.theme-max #tab-settings input[type="text"],
    body.theme-max #tab-settings input[type="number"],
    body.theme-max #tab-settings input[type="password"],
    body.theme-touch #tab-controls input[type="text"],
    body.theme-touch #tab-controls input[type="number"],
    body.theme-touch #tab-controls input[type="password"],
    body.theme-touch #tab-settings input[type="text"],
    body.theme-touch #tab-settings input[type="number"],
    body.theme-touch #tab-settings input[type="password"],
    body.theme-thumb #tab-controls input[type="text"],
    body.theme-thumb #tab-controls input[type="number"],
    body.theme-thumb #tab-controls input[type="password"],
    body.theme-thumb #tab-settings input[type="text"],
    body.theme-thumb #tab-settings input[type="number"],
    body.theme-thumb #tab-settings input[type="password"] {
      border-color: var(--ui-border);
      border-radius: var(--ui-radius);
      color: var(--ui-text);
      background: var(--ui-section-muted-bg, var(--ui-section-bg));
    }
    body.theme-original #tab-settings select,
    body.theme-original #tab-settings select.input-field,
    body.theme-dark #tab-settings select,
    body.theme-dark #tab-settings select.input-field,
    body.theme-max #tab-settings select,
    body.theme-max #tab-settings select.input-field,
    body.theme-touch #tab-settings select,
    body.theme-touch #tab-settings select.input-field,
    body.theme-thumb #tab-settings select,
    body.theme-thumb #tab-settings select.input-field {
      border-color: var(--ui-border);
      border-radius: var(--ui-radius);
      color: var(--ui-text);
      background: var(--ui-section-muted-bg, var(--ui-section-bg));
    }
    body.theme-dark .web-preset-empty-link,
    body.theme-thumb .web-preset-empty-link {
      color: var(--ui-accent);
    }
    body.theme-dark .web-preset-launch-label,
    body.theme-thumb .web-preset-launch-label {
      color: var(--ui-muted);
    }
    @media (min-width: 640px) {
      body.theme-original #tab-settings .preset-group:has(> label),
      body.theme-dark #tab-settings .preset-group:has(> label),
      body.theme-max #tab-settings .preset-group:has(> label),
      body.theme-touch #tab-settings .preset-group:has(> label),
      body.theme-thumb #tab-settings .preset-group:has(> label) {
        display: grid;
        grid-template-columns: minmax(120px, 180px) minmax(0, 1fr);
        gap: 8px 16px;
        align-items: start;
      }
      body.theme-original #tab-settings .preset-group:has(> label) > label,
      body.theme-dark #tab-settings .preset-group:has(> label) > label,
      body.theme-max #tab-settings .preset-group:has(> label) > label,
      body.theme-touch #tab-settings .preset-group:has(> label) > label,
      body.theme-thumb #tab-settings .preset-group:has(> label) > label {
        grid-column: 1;
        margin-bottom: 0;
        padding-top: 10px;
      }
      body.theme-original #tab-settings .preset-group:has(> label) > *:not(label),
      body.theme-dark #tab-settings .preset-group:has(> label) > *:not(label),
      body.theme-max #tab-settings .preset-group:has(> label) > *:not(label),
      body.theme-touch #tab-settings .preset-group:has(> label) > *:not(label),
      body.theme-thumb #tab-settings .preset-group:has(> label) > *:not(label) {
        grid-column: 2;
        min-width: 0;
      }
      body.theme-original #tab-settings .preset-group:has(> label) > small,
      body.theme-dark #tab-settings .preset-group:has(> label) > small,
      body.theme-max #tab-settings .preset-group:has(> label) > small,
      body.theme-touch #tab-settings .preset-group:has(> label) > small,
      body.theme-thumb #tab-settings .preset-group:has(> label) > small {
        grid-column: 1 / -1;
        padding-top: 2px;
      }
      body.theme-original #tab-settings .preset-group:has(> label) > #web-backup-ip-list,
      body.theme-dark #tab-settings .preset-group:has(> label) > #web-backup-ip-list,
      body.theme-max #tab-settings .preset-group:has(> label) > #web-backup-ip-list,
      body.theme-touch #tab-settings .preset-group:has(> label) > #web-backup-ip-list,
      body.theme-thumb #tab-settings .preset-group:has(> label) > #web-backup-ip-list {
        grid-column: 1 / -1;
      }
      body.theme-original #tab-settings .preset-group:has(> label) > button.btn,
      body.theme-dark #tab-settings .preset-group:has(> label) > button.btn,
      body.theme-max #tab-settings .preset-group:has(> label) > button.btn,
      body.theme-touch #tab-settings .preset-group:has(> label) > button.btn,
      body.theme-thumb #tab-settings .preset-group:has(> label) > button.btn {
        grid-column: 1 / -1;
      }
    }
    body.theme-original #tab-settings input[type="checkbox"] + label,
    body.theme-dark #tab-settings input[type="checkbox"] + label,
    body.theme-max #tab-settings input[type="checkbox"] + label,
    body.theme-touch #tab-settings input[type="checkbox"] + label,
    body.theme-thumb #tab-settings input[type="checkbox"] + label {
      font-weight: 400 !important;
      margin-bottom: 0 !important;
      color: var(--ui-text);
    }
  </style>
  ${webUiCustomCssPath ? '<link rel="stylesheet" href="/custom-style.css?v=' + Date.now() + '">' : ''}
</head>
<body class="theme-${webUiTheme}" data-theme="${webUiTheme}">
  <div class="container">
    <div class="web-ui-header">
      <h1>
      ${showLogo ? '<img class="web-ui-brand-logo" src="/custom-logo?v=' + Date.now() + '" alt="">' : '<svg class="system-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="4" width="20" height="12" rx="2" ry="2"></rect><line x1="6" y1="20" x2="18" y2="20"></line><line x1="8" y1="16" x2="8" y2="20"></line><line x1="16" y1="16" x2="16" y2="20"></line><circle cx="12" cy="10" r="3" fill="currentColor"></circle><polygon points="10 10 12 9 14 10 12 11" fill="white"></polygon></svg>'}
      ${machineName}
    </h1>
    </div>
    
    <!-- Tabs -->
    <div class="tabs">
      <button class="tab-btn active" data-tab="remote">Remote</button>
      <button class="tab-btn" data-tab="controls">Controls</button>
      ${!webUiRestrictedTunnelClient ? '<button class="tab-btn" data-tab="settings">Settings</button>' : ''}
    </div>
    
    <!-- Remote Tab (Default) -->
    <div id="tab-remote" class="tab-content active">
      <div class="remote-header-compact">
        <span class="remote-status-dot" aria-hidden="true"></span>
        <span class="remote-machine-name">${machineName}</span>
        <span class="slide-counter" id="remote-slide-counter">—</span>
        <button type="button" class="notes-toggle-btn" id="notes-toggle-btn" title="Toggle speaker notes" aria-label="Toggle speaker notes">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
          </svg>
          <span class="toggle-btn-text">Notes</span>
        </button>
        <button type="button" class="preview-toggle-btn" id="previews-toggle-btn" title="Toggle slide previews" aria-label="Toggle slide previews">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="3" y="4" width="8" height="8" rx="1"></rect>
            <rect x="13" y="4" width="8" height="8" rx="1"></rect>
            <rect x="3" y="14" width="18" height="6" rx="1"></rect>
          </svg>
          <span class="toggle-btn-text">Previews</span>
        </button>
      </div>
      <div class="stagetimer-container disabled" id="stagetimer-container" style="display: none;">
        <div class="stagetimer-row">
          <span class="stagetimer-dot" aria-hidden="true"></span>
          <div class="stagetimer-label" id="stagetimer-label">Stage Timer</div>
          <div class="stagetimer-time" id="stagetimer-time">--:--</div>
        </div>
        <div class="stagetimer-status" id="stagetimer-status">Not configured</div>
        <div class="stagetimer-messages" id="stagetimer-messages"></div>
      </div>
      <div class="slide-previews-container" id="slide-previews-container">
        <div class="slide-previews-grid">
          <div class="slide-preview-card clickable" id="slide-preview-current-card" title="Click to go to previous slide">
            <img class="slide-preview-img empty" id="slide-preview-current-img" alt="Current slide — click to go back" src="data:image/gif;base64,R0lGODlhAQABAAD/ACwAAAAAAQABAAACADs=" />
            <div class="slide-preview-label" id="slide-preview-current-label">Current Slide</div>
          </div>
          <div class="slide-preview-card clickable" id="slide-preview-next-card" title="Click to go to next slide">
            <img class="slide-preview-img empty" id="slide-preview-next-img" alt="Next slide — click to advance" src="data:image/gif;base64,R0lGODlhAQABAAD/ACwAAAAAAQABAAACADs=" />
            <div class="slide-preview-label" id="slide-preview-next-label">Next Slide</div>
          </div>
        </div>
      </div>

      <div class="speaker-notes-container" id="speaker-notes-container">
        <div class="notes-zoom-controls" id="notes-zoom-controls">
          <span class="speaker-notes-toolbar-label" id="speaker-notes-toolbar-label">Speaker notes · <span id="speaker-notes-slide-num">—</span></span>
          <span class="notes-zoom-toolbar-actions">
            <button type="button" class="notes-zoom-btn" id="btn-scroll-notes-up" title="Scroll speaker notes up" aria-label="Scroll speaker notes up">
              <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><polyline points="18 15 12 9 6 15"></polyline></svg>
            </button>
            <button type="button" class="notes-zoom-btn" id="btn-scroll-notes-down" title="Scroll speaker notes down" aria-label="Scroll speaker notes down">
              <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"></polyline></svg>
            </button>
            <button type="button" class="notes-zoom-btn" id="notes-zoom-out" aria-label="Zoom out">−</button>
            <span id="notes-zoom-readout">18px</span>
            <button type="button" class="notes-zoom-btn" id="notes-zoom-in" aria-label="Zoom in">+</button>
          </span>
        </div>
        <div class="speaker-notes-content-wrapper">
          <div class="speaker-notes-content" id="speaker-notes-content">Loading notes...</div>
          <div id="notes-encoding-warning" style="display:none; margin-top:6px; padding:6px 10px; font-size:12px; line-height:1.4; color:#b26a00; background:rgba(255,193,7,0.12); border:1px solid rgba(255,193,7,0.3); border-radius:6px;">Line break encoding issues detected on this slide. Notes are displayed with corrections applied. To fix permanently, re-enter line breaks in the Google Slides editor or run the <a href="https://github.com/TomsFaire/Google-Slides-Controller/blob/main/docs/fix-speaker-notes.gs" target="_blank" rel="noopener" style="color:inherit;text-decoration:underline;">cleanup script</a>.</div>
        </div>
      </div>
      <div class="remote-controls" id="remote-controls">
        <button type="button" class="remote-btn remote-btn-prev" id="remote-btn-prev">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3">
            <polyline points="15 18 9 12 15 6"></polyline>
          </svg>
          <span class="remote-btn-label">Previous</span>
        </button>
        <button type="button" class="remote-btn remote-btn-next" id="remote-btn-next">
          <span class="remote-btn-label">Next slide</span>
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3">
            <polyline points="9 18 15 12 9 6"></polyline>
          </svg>
        </button>
      </div>
    </div>
    
    <!-- Controls Tab -->
    <div id="tab-controls" class="tab-content">
      <div class="info">
        Use these controls to manage your active presentation.
      </div>
      
      <!-- Open Presentation -->
      <div class="controls-section">
        <h3>Open Presentation</h3>
        <div class="preset-group">
          <label for="presentation-url">Google Slides URL</label>
          <input type="text" id="presentation-url" name="presentation-url" placeholder="https://docs.google.com/presentation/d/..." />
        </div>
        <div style="display: flex; gap: 10px;">
          <button type="button" class="btn" id="btn-open-presentation" style="flex: 1;">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;">
              <polyline points="5 12 3 12 12 3 21 12 19 12"></polyline>
              <path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path>
              <polyline points="9 21 9 12 15 12 15 21"></polyline>
            </svg>
            Launch Presentation
          </button>
          <button type="button" class="btn" id="btn-open-presentation-with-notes" style="flex: 1;">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;">
              <polyline points="5 12 3 12 12 3 21 12 19 12"></polyline>
              <path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path>
              <polyline points="9 21 9 12 15 12 15 21"></polyline>
            </svg>
            Launch with Notes
          </button>
        </div>
      </div>
      
      <!-- Preset Presentations -->
      <div class="controls-section">
        <h3>Preset Presentations</h3>
        <div id="preset-buttons-container" style="display: flex; flex-direction: column; gap: 10px;">
          <!-- Preset buttons will be dynamically loaded here -->
        </div>
      </div>
      
      <!-- Speaker Notes Controls -->
      <div class="controls-section">
        <h3>Speaker Notes</h3>
        <button type="button" class="btn-control" id="btn-start-notes" title="Start speaker notes window">
          <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="16" y1="13" x2="8" y2="13"></line>
            <line x1="16" y1="17" x2="8" y2="17"></line>
            <polyline points="10 9 9 9 8 9"></polyline>
          </svg>
          Start Notes
        </button>
      </div>
      
      <!-- Presentation Controls -->
      <div class="controls-section">
        <h3>Presentation Controls</h3>
        <div class="controls-grid">
          <button type="button" class="btn-control" id="btn-prev-slide" data-tooltip="Go to previous slide">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="15 18 9 12 15 6"></polyline>
            </svg>
            Previous Slide
          </button>
          <button type="button" class="btn-control" id="btn-next-slide" data-tooltip="Go to next slide">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="9 18 15 12 9 6"></polyline>
            </svg>
            Next Slide
          </button>
          <button type="button" class="btn-control" id="btn-reload" data-tooltip="Reload presentation and return to current slide">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="23 4 23 10 17 10"></polyline>
              <polyline points="1 20 1 14 7 14"></polyline>
              <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path>
            </svg>
            Reload Presentation
          </button>
          <button type="button" class="btn-control" id="btn-close-presentation" data-tooltip="Close current presentation">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <line x1="18" y1="6" x2="6" y2="18"></line>
              <line x1="6" y1="6" x2="18" y2="18"></line>
            </svg>
            Close Presentation
          </button>
        </div>
      </div>
    </div>
    
    ${!webUiRestrictedTunnelClient ? `<!-- Settings Tab (Hidden by default) -->
    <div id="tab-settings" class="tab-content">
      <!-- Monitor Setup Section -->
      <div class="controls-section">
        <h3>Monitor Setup</h3>
        <div class="info" style="margin-bottom: 15px;">
          Select which monitors to use for the presentation and speaker notes windows.
        </div>
        <div class="preset-group">
          <label for="web-presentation-display">Presentation Monitor</label>
          <select id="web-presentation-display" class="input-field" style="width: 100%; padding: 8px;">
            <option value="">Loading displays...</option>
          </select>
        </div>
        <div class="preset-group">
          <label for="web-notes-display">Notes Monitor</label>
          <select id="web-notes-display" class="input-field" style="width: 100%; padding: 8px;">
            <option value="">Loading displays...</option>
          </select>
        </div>
        <div class="preset-group" style="margin-top: 10px;">
          <label for="web-notes-layout">Notes Layout</label>
          <select id="web-notes-layout" class="input-field" style="width: 100%; padding: 8px;">
            <option value="hide">Full Notes (slide previews hidden)</option>
            <option value="default">Google Default (50/50 split)</option>
          </select>
          <small style="display: block; margin-top: 5px; color: #888; font-size: 12px;">Applies on next notes launch. Use Relaunch Notes to apply immediately.</small>
        </div>
        <div class="preset-group" style="margin-top: 6px;">
          <button type="button" class="btn" id="btn-relaunch-notes" style="width: 100%;">Relaunch Notes</button>
        </div>
        <div class="preset-group" style="margin-top: 10px;">
          <label for="web-default-notes-zoom-steps">Default speaker notes zoom (steps)</label>
          <input type="number" id="web-default-notes-zoom-steps" class="input-field" style="width: 100%; padding: 8px;" min="-10" max="40" value="0" step="1" />
          <small style="display: block; margin-top: 5px; color: #888; font-size: 12px;">Extra Zoom in clicks when notes open (0 = Slides default). Applies on next presentation or when notes reopen.</small>
        </div>
        <button type="button" class="btn" id="btn-save-displays" style="margin-top: 10px;">Save Monitor Settings</button>
      </div>
      
      <!-- Machine Name Section -->
      <div class="controls-section">
        <h3>Machine Name</h3>
        <div class="info" style="margin-bottom: 15px;">
          Set a name for this machine (shown in web UI header).
        </div>
        <div class="preset-group">
          <label for="web-machine-name">Machine Name</label>
          <input type="text" id="web-machine-name" class="input-field" placeholder="Enter machine name..." maxlength="50" />
          <small style="display: block; margin-top: 5px; color: #888; font-size: 12px;">Leave empty to use system hostname</small>
        </div>
        <button type="button" class="btn" id="btn-save-machine-name" style="margin-top: 10px;">Save Machine Name</button>
      </div>
      
      <!-- Primary/Backup Configuration Section -->
      <div class="controls-section">
        <h3>Primary/Backup Configuration</h3>
        <div class="info" style="margin-bottom: 15px;">
          Configure this instance as primary (controls backups) or backup (follows primary).
        </div>
        <div class="preset-group">
          <label>Mode</label>
          <div style="display: flex; gap: 20px; margin-top: 8px;">
            <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
              <input type="radio" name="web-primary-backup-mode" id="web-mode-primary" value="primary" />
              <span>Primary</span>
            </label>
            <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
              <input type="radio" name="web-primary-backup-mode" id="web-mode-backup" value="backup" />
              <span>Backup</span>
            </label>
            <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
              <input type="radio" name="web-primary-backup-mode" id="web-mode-standalone" value="standalone" checked />
              <span>Standalone</span>
            </label>
          </div>
          <small style="display: block; margin-top: 5px; color: #888; font-size: 12px;">Primary: Controls backup machines. Backup: Follows primary commands. Standalone: Independent operation.</small>
        </div>
        
        <div id="web-backup-config" style="display: none; margin-top: 15px;">
          <div class="preset-group">
            <label for="web-backup-port">Backup Communication Port</label>
            <input type="number" id="web-backup-port" class="input-field" min="1024" max="65535" placeholder="9595" />
            <small style="display: block; margin-top: 5px; color: #888; font-size: 12px;">Port used to communicate with backup machines (default: 9595)</small>
          </div>
          
          <div class="preset-group">
            <label>Backup Machines</label>
            <div id="web-backup-ip-list" style="display: flex; flex-direction: column; gap: 10px; width: 100%; min-width: 0;"></div>
            <button type="button" class="btn btn-secondary" id="web-add-backup-ip" style="margin-top: 10px;">+ Add backup machine</button>
            <small style="display: block; margin-top: 5px; color: #888; font-size: 12px;">Enter an IP address or hostname for each backup. Supports any number of backups.</small>
          </div>
        </div>
        
        <button type="button" class="btn" id="btn-save-primary-backup" style="margin-top: 15px;">Save Primary/Backup Settings</button>
      </div>
      
      <!-- Network Ports Section -->
      <div class="controls-section">
        <h3>Network Ports</h3>
        <div class="info" style="margin-bottom: 15px;">
          Configure ports for API and Web UI (restart required for changes to take effect).
        </div>
        <div class="preset-group">
          <label for="web-api-port">API Port (Companion)</label>
          <input type="number" id="web-api-port" class="input-field" min="1024" max="65535" placeholder="9595" />
          <small style="display: block; margin-top: 5px; color: #888; font-size: 12px;">Port for Companion module API (default: 9595)</small>
        </div>
        <div class="preset-group">
          <label for="web-web-ui-port">Web UI Port</label>
          <input type="number" id="web-web-ui-port" class="input-field" min="1" max="65535" placeholder="80" />
          <small style="display: block; margin-top: 5px; color: #888; font-size: 12px;">Port for web interface (default: 80, requires admin for ports &lt;1024)</small>
        </div>
        <button type="button" class="btn" id="btn-save-ports" style="margin-top: 10px;">Save Port Settings</button>
        <div class="web-ui-callout web-ui-callout--warning">
          ⚠️ Port changes require restarting the app to take effect.
        </div>
      </div>
      
      <!-- Preset Presentations Section -->
      <div class="controls-section" id="web-preset-section">
        <h3>Preset Presentations</h3>
        <div class="info" style="margin-bottom: 15px;">
          Configure preset presentations. These can be opened from Companion or the Remote tab. Preset 1, 2, 3… correspond to Companion actions.
        </div>
      
      <form id="preset-form">
      <div id="web-preset-list" style="display: flex; flex-direction: column; gap: 10px; margin-bottom: 10px;"></div>
      <button type="button" class="btn btn-secondary" id="web-add-preset" style="margin-bottom: 12px;">+ Add presentation</button>
        <button type="submit" class="btn">Save Presets</button>
        <button type="button" class="btn btn-secondary" id="load-btn">Load Current Presets</button>
      </form>
      </div>
      
      <!-- Stagetimer Integration -->
      <div class="controls-section">
        <h3>Stagetimer.io Integration</h3>
        <div class="info" style="margin-bottom: 20px;">
          Connect to your stagetimer.io room to display live timer data. Get your Room ID and API Key from the stagetimer.io controller page.
        </div>
        <div class="preset-group">
          <label for="stagetimer-room-id">Room ID</label>
          <input type="text" id="stagetimer-room-id" name="stagetimer-room-id" placeholder="Enter your stagetimer.io Room ID" />
        </div>
        <div class="preset-group">
          <label for="stagetimer-api-key">API Key</label>
          <input type="password" id="stagetimer-api-key" name="stagetimer-api-key" placeholder="Enter your stagetimer.io API Key" />
        </div>
        <div style="display: flex; align-items: center; gap: 10px; margin-top: 12px;">
          <input type="checkbox" id="stagetimer-enabled" style="width: auto;" />
          <label for="stagetimer-enabled" style="margin: 0; font-weight: normal;">Enable timer display</label>
        </div>
        <div style="display: flex; align-items: center; gap: 10px; margin-top: 12px;">
          <input type="checkbox" id="stagetimer-visible" style="width: auto;" checked />
          <label for="stagetimer-visible" style="margin: 0; font-weight: normal;">Show timer on Remote tab</label>
        </div>
        <button type="button" class="btn" id="btn-save-stagetimer" style="margin-top: 12px;">Save Stagetimer Settings</button>
        <button type="button" class="btn btn-secondary" id="btn-load-stagetimer" style="margin-top: 8px;">Load Current Settings</button>
      </div>

      <!-- WAN Tunnel Section -->
      <div class="controls-section">
        <h3>WAN Access (Cloudflare Tunnel)</h3>
        <div id="web-tunnel-status" class="info" style="margin-bottom: 12px;">Checking tunnel status…</div>
        <div id="web-tunnel-qr-row" style="display: none; align-items: center; gap: 8px; flex-wrap: wrap; margin-top: 8px;">
          <input type="number" id="web-tunnel-qr-duration" value="20" min="5" max="300"
            style="width: 70px; padding: 6px 8px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px;" />
          <label style="margin: 0; font-weight: normal; font-size: 13px;">sec</label>
          <button type="button" class="btn" id="btn-show-tunnel-qr" style="margin: 0;">Show QR on Notes Display</button>
          <button type="button" class="btn btn-secondary" id="btn-hide-tunnel-qr" style="margin: 0;">Hide QR</button>
        </div>
        <small style="display: block; margin-top: 8px; color: #888; font-size: 12px;">
          Shows a scannable QR code on the presenter's notes monitor. Enable the tunnel from the desktop app Settings.
        </small>
      </div>

      <!-- Logging Section -->
      <div class="controls-section">
        <h3>Logging</h3>
        <div class="info" style="margin-bottom: 10px;">
          Control how much the app writes to its terminal logs.
        </div>
        <div style="display: flex; align-items: center; gap: 10px; margin-top: 12px;">
          <input type="checkbox" id="web-verbose-logging" style="width: auto;" />
          <label for="web-verbose-logging" style="margin: 0; font-weight: normal;">Enable verbose logging</label>
        </div>
        <small style="display: block; margin-top: 6px; color: #888; font-size: 12px;">
          Verbose logs help debugging. Secrets (API keys/tokens/passwords) are always redacted from logs.
        </small>
        <button type="button" class="btn" id="btn-save-logging" style="margin-top: 12px;">Save Logging Settings</button>
      </div>
      
      ${webUiDebugConsoleEnabled ? `
      <!-- Debug Console (enabled from desktop app) -->
      <div class="controls-section">
        <h3>Debug Console</h3>
        <div class="info" style="margin-bottom: 10px;">
          Console output for debugging stagetimer integration and other issues.
        </div>
        <div style="background: #1e1e1e; color: #d4d4d4; padding: 12px; border-radius: 8px; font-family: 'Courier New', monospace; font-size: 12px; max-height: 300px; overflow-y: auto; margin-bottom: 10px;" id="debug-console">
          <div style="color: #888;">Console ready. Logs will appear here...</div>
        </div>
        <button type="button" class="btn btn-secondary" id="btn-clear-console" style="margin-top: 8px;">Clear Console</button>
      </div>
      ` : ``}
    </div>
` : ''}

    <div id="status" class="status"></div>
    <div class="build-number">${versionString}</div>
  </div>

  <!-- Outside .container: fixed bottom nav is clipped if left inside overflow:hidden .container (Chrome) -->
  <nav class="bottom-tabs" aria-label="Main tabs">
    <button type="button" class="tab-btn active" data-tab="remote">
      <svg class="tab-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="2" y="4" width="20" height="12" rx="2"></rect><path d="M6 20h12M8 16v4M16 16v4"></path></svg>
      <span>Remote</span>
    </button>
    <button type="button" class="tab-btn" data-tab="controls">
      <svg class="tab-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="3"></circle><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"></path></svg>
      <span>Controls</span>
    </button>
    ${!webUiRestrictedTunnelClient ? `<button type="button" class="tab-btn" data-tab="settings">
      <svg class="tab-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>
      <span>Settings</span>
    </button>` : ''}
  </nav>
  
  <script>
    const form = document.getElementById('preset-form');
    const loadBtn = document.getElementById('load-btn');
    const status = document.getElementById('status');
    // Use relative URLs so API calls go through the Web UI server (port 80)
    // The Web UI server will proxy these requests to the API server (port 9595)
    // This allows the Web UI to work even when only port 80 is accessible from the network
    const API_BASE = '';
    window.__GSO_WEB_UI_RESTRICTED__ = ${webUiRestrictedTunnelClient ? 'true' : 'false'};
    
    // Debug: Log the API base URL for troubleshooting
    console.log('[Web UI] Using relative API URLs (proxied through Web UI server on port 80)');
    console.log('[Web UI] window.location.hostname:', window.location.hostname);
    console.log('[Web UI] window.location.host:', window.location.host);
    
    function showStatus(message, isError) {
      status.textContent = message;
      status.className = 'status ' + (isError ? 'error' : 'success');
      setTimeout(() => {
        status.className = 'status';
        status.textContent = ''; // Clear text when hidden
      }, 3000);
    }
    
    // Prevent native tooltips and use custom floating ones (skip prev/next slide buttons)
    document.querySelectorAll('.btn-control[title]').forEach(btn => {
      // Skip prev/next slide buttons - no tooltips for those
      if (btn.id === 'btn-prev-slide' || btn.id === 'btn-next-slide') {
        btn.removeAttribute('title');
        return;
      }
      
      const titleText = btn.getAttribute('title');
      btn.setAttribute('data-tooltip', titleText);
      btn.removeAttribute('title'); // Remove native title to prevent layout shift
      
      // Restore title for accessibility when not hovering
      btn.addEventListener('mouseenter', function() {
        this.removeAttribute('title');
      });
      btn.addEventListener('mouseleave', function() {
        this.setAttribute('title', titleText);
      });
    });
    
    // Tab switching (sync all .tab-btn duplicates, e.g. top bar + bottom bar on theme-light)
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const tabName = btn.getAttribute('data-tab');
        if (window.__GSO_WEB_UI_RESTRICTED__ && tabName === 'settings') return;
        const tabPanel = document.getElementById('tab-' + tabName);
        if (!tabPanel) return;

        document.querySelectorAll('.tab-btn').forEach(b => {
          b.classList.toggle('active', b.getAttribute('data-tab') === tabName);
        });

        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        tabPanel.classList.add('active');
      });
    });

    document.getElementById('preset-buttons-container').addEventListener('click', (e) => {
      const link = e.target.closest('.web-preset-empty-link');
      if (!link) return;
      e.preventDefault();
      if (window.__GSO_WEB_UI_RESTRICTED__) return;
      const settingsBtn = document.querySelector('.tab-btn[data-tab="settings"]');
      if (!settingsBtn) return;
      settingsBtn.click();
      requestAnimationFrame(() => {
        document.getElementById('web-preset-section')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      });
    });

    function apiCall(endpoint, method = 'POST') {
      const url = API_BASE + endpoint;
      console.log('[Web UI] Making API call:', method, url);
      
      return fetch(url, {
        method: method,
        headers: { 'Content-Type': 'application/json' }
      })
        .then(res => {
          if (!res.ok) {
            throw new Error('HTTP error! status: ' + res.status);
          }
          return res.json();
        })
        .then(result => {
          console.log('[Web UI] API response:', result);
          if (result.success !== false) {
            showStatus(result.message || 'Action completed successfully', false);
          } else {
            showStatus(result.error || 'Action failed', true);
          }
          return result;
        })
        .catch(err => {
          console.error('[Web UI] API call error:', err);
          console.error('[Web UI] Failed URL:', url);
          let errorMsg = 'Failed: ' + err.message;
          if (err.message.includes('Failed to fetch') || err.message.includes('NetworkError')) {
            errorMsg += ' (Cannot reach API server at ' + API_BASE + '. Check network connection and firewall settings.)';
          } else {
            errorMsg += ' (Make sure the app is running)';
          }
          showStatus(errorMsg, true);
          throw err;
        });
    }
    
    // Haptic feedback function for mobile devices
    function triggerHapticFeedback() {
      if ('vibrate' in navigator) {
        // Light vibration for button press
        navigator.vibrate(10);
      }
    }
    
    // Set up control buttons
    document.getElementById('btn-prev-slide').addEventListener('click', () => {
      apiCall('/api/previous-slide').then(() => {
        updateSlideButtons();
      });
    });
    
    document.getElementById('btn-next-slide').addEventListener('click', () => {
      apiCall('/api/next-slide').then(() => {
        updateSlideButtons();
      });
    });
    
    document.getElementById('btn-reload').addEventListener('click', () => {
      apiCall('/api/reload-presentation').then(() => {
        updateSlideButtons();
      });
    });
    
    document.getElementById('btn-close-presentation').addEventListener('click', () => {
      apiCall('/api/close-presentation').then(() => {
        updateSlideButtons();
      });
    });
    
    // Remote tab buttons
    // Speaker notes functionality
    let notesVisible = false;
    let notesZoomLevel = 1; // Numeric zoom level (1 = normal, can go up/down continuously)
    let previewsVisible = false;
    
    function normalizeSpeakerNotes(text) {
      if (text == null) return '';
      var s = String(text);
      var nl = String.fromCharCode(10);
      s = s.replace(/\\r\\n/g, nl).replace(/\\r/g, nl).replace(/\\u2028/g, nl).replace(/\\u2029/g, nl).replace(/\\uFFFD+/g, nl).replace(/\\u0000/g, '');
      return s;
    }

    function updateRemoteSlideContext(current, total) {
      const counter = document.getElementById('remote-slide-counter');
      const notesNum = document.getElementById('speaker-notes-slide-num');
      var counterText = '—';
      var slideNumText = '—';
      if (typeof current === 'number' && !isNaN(current)) {
        slideNumText = String(current);
        if (typeof total === 'number' && !isNaN(total)) {
          counterText = current + ' / ' + total;
        } else {
          counterText = String(current);
        }
      }
      if (counter) counter.textContent = counterText;
      if (notesNum) notesNum.textContent = slideNumText;
    }

    function loadSpeakerNotes() {
      fetch(API_BASE + '/api/get-speaker-notes')
        .then(res => res.json())
        .then(data => {
          const notesContent = document.getElementById('speaker-notes-content');
          var raw = (data.success && data.notes) ? data.notes : (data.notes || 'No notes available. Make sure speaker notes are open.');
          notesContent.textContent = normalizeSpeakerNotes(raw);
          var warn = document.getElementById('notes-encoding-warning');
          if (warn) warn.style.display = data.encodingIssuesDetected ? 'block' : 'none';
        })
        .catch(err => {
          console.error('Failed to load speaker notes:', err);
          document.getElementById('speaker-notes-content').textContent = 'Failed to load notes.';
          var warn = document.getElementById('notes-encoding-warning');
          if (warn) warn.style.display = 'none';
        });
    }

    function loadSlidePreviews() {
      fetch(API_BASE + '/api/get-slide-previews')
        .then(res => res.json())
        .then(data => {
          const currentImg = document.getElementById('slide-preview-current-img');
          const nextImg = document.getElementById('slide-preview-next-img');
          const currentLabel = document.getElementById('slide-preview-current-label');
          const nextLabel = document.getElementById('slide-preview-next-label');
          const placeholderSrc = 'data:image/gif;base64,R0lGODlhAQABAAD/ACwAAAAAAQABAAACADs=';

          if (!data || !data.success) {
            const msg = (data && data.error) ? data.error : 'Previews unavailable';
            currentLabel.textContent = 'Current Slide (preview unavailable)';
            nextLabel.textContent = 'Next Slide (preview unavailable)';
            if (currentImg) {
              currentImg.src = placeholderSrc;
              currentImg.classList.add('empty');
            }
            if (nextImg) {
              nextImg.src = placeholderSrc;
              nextImg.classList.add('empty');
            }
            console.debug('[Web UI] Slide previews unavailable:', msg);
            return;
          }

          const curNum = data.currentSlide;
          const nextNum = data.nextSlide;

          currentLabel.textContent = (curNum ? ('Current Slide (' + curNum + ')') : 'Current Slide');
          nextLabel.textContent = (nextNum ? ('Next Slide (' + nextNum + ')') : 'Next Slide');

          updateRemoteSlideContext(data.currentSlide, data.totalSlides);

          if (currentImg && data.current && typeof data.current.dataUrl === 'string' && data.current.dataUrl.startsWith('data:image/')) {
            currentImg.src = data.current.dataUrl;
            currentImg.classList.remove('empty');
          }
          if (nextImg && data.next && typeof data.next.dataUrl === 'string' && data.next.dataUrl.startsWith('data:image/')) {
            nextImg.src = data.next.dataUrl;
            nextImg.classList.remove('empty');
          }
        })
        .catch(err => {
          console.debug('[Web UI] Failed to load slide previews:', err.message);
        });
    }

    function closeNotesUi() {
      const btn = document.getElementById('notes-toggle-btn');
      const container = document.getElementById('speaker-notes-container');
      const controls = document.getElementById('remote-controls');
      const zoomControls = document.getElementById('notes-zoom-controls');
      const body = document.body;

      notesVisible = false;
      if (btn) btn.classList.remove('active');
      if (container) container.classList.remove('visible');
      if (zoomControls) zoomControls.classList.remove('visible');
      if (controls) controls.classList.remove('with-notes');
      body.classList.remove('notes-visible');

      if (window.notesRefreshInterval) {
        clearInterval(window.notesRefreshInterval);
        window.notesRefreshInterval = null;
      }
    }

    function closePreviewsUi() {
      const btn = document.getElementById('previews-toggle-btn');
      const container = document.getElementById('slide-previews-container');
      const controls = document.getElementById('remote-controls');
      const body = document.body;

      previewsVisible = false;
      if (btn) btn.classList.remove('active');
      if (container) container.classList.remove('visible');
      if (controls) controls.classList.remove('with-panel');
      body.classList.remove('previews-visible');

      if (window.previewsRefreshInterval) {
        clearInterval(window.previewsRefreshInterval);
        window.previewsRefreshInterval = null;
      }
    }
    
    function updateNotesZoomReadout() {
      const notesContent = document.getElementById('speaker-notes-content');
      const readout = document.getElementById('notes-zoom-readout');
      if (readout && notesContent) {
        const px = parseInt(window.getComputedStyle(notesContent).fontSize, 10);
        readout.textContent = (isNaN(px) ? 18 : px) + 'px';
      }
    }

    function updateNotesZoom() {
      const notesContent = document.getElementById('speaker-notes-content');
      // Calculate font size based on zoom level (18px base, +/- 2px per level)
      const baseSize = 18;
      const fontSize = baseSize + ((notesZoomLevel - 1) * 2);
      notesContent.style.fontSize = fontSize + 'px';
      updateNotesZoomReadout();
    }
    
    document.getElementById('notes-toggle-btn').addEventListener('click', () => {
      const btn = document.getElementById('notes-toggle-btn');
      const container = document.getElementById('speaker-notes-container');
      const controls = document.getElementById('remote-controls');
      const zoomControls = document.getElementById('notes-zoom-controls');
      const body = document.body;
      
      if (!notesVisible) {
        // Opening notes - first check if speaker notes window is open, if not, open it
        fetch(API_BASE + '/api/get-speaker-notes')
          .then(res => res.json())
          .then(data => {
            // If notes window is not open, open it first
            if (!data.success && data.error && data.error.includes('No speaker notes window')) {
              console.log('[Web UI] Speaker notes not open, opening them first...');
              return apiCall('/api/open-speaker-notes').then(() => {
                // Wait a moment for notes to open, then show the UI
                setTimeout(() => {
                  notesVisible = true;
                  btn.classList.add('active');
                  container.classList.add('visible');
                  controls.classList.add('with-notes');
                  // If previews are already open, keep compact layout class too
                  if (previewsVisible) controls.classList.add('with-panel');
                  zoomControls.classList.add('visible');
                  body.classList.add('notes-visible');
                  loadSpeakerNotes();
                  // Refresh notes every 2 seconds when visible
                  if (window.notesRefreshInterval) clearInterval(window.notesRefreshInterval);
                  window.notesRefreshInterval = setInterval(loadSpeakerNotes, 2000);
                }, 1000);
              });
            } else {
              // Notes are already open, just show the UI
              notesVisible = true;
              btn.classList.add('active');
              container.classList.add('visible');
              controls.classList.add('with-notes');
              if (previewsVisible) controls.classList.add('with-panel');
              zoomControls.classList.add('visible');
              body.classList.add('notes-visible');
              loadSpeakerNotes();
              // Refresh notes every 2 seconds when visible
              if (window.notesRefreshInterval) clearInterval(window.notesRefreshInterval);
              window.notesRefreshInterval = setInterval(loadSpeakerNotes, 2000);
            }
          })
          .catch(err => {
            console.error('[Web UI] Error checking/opening speaker notes:', err);
            // Try to open notes anyway
            apiCall('/api/open-speaker-notes').then(() => {
              setTimeout(() => {
                notesVisible = true;
                btn.classList.add('active');
                container.classList.add('visible');
                controls.classList.add('with-notes');
                if (previewsVisible) controls.classList.add('with-panel');
                zoomControls.classList.add('visible');
                body.classList.add('notes-visible');
                loadSpeakerNotes();
                if (window.notesRefreshInterval) clearInterval(window.notesRefreshInterval);
                window.notesRefreshInterval = setInterval(loadSpeakerNotes, 2000);
              }, 1000);
            });
          });
      } else {
        // Closing notes
        closeNotesUi();
      }
    });

    document.getElementById('previews-toggle-btn').addEventListener('click', () => {
      const btn = document.getElementById('previews-toggle-btn');
      const container = document.getElementById('slide-previews-container');
      const controls = document.getElementById('remote-controls');
      const body = document.body;

      if (!previewsVisible) {
        // Opening previews - requires the speaker notes window (Presenter View) to be open
        fetch(API_BASE + '/api/get-slide-previews')
          .then(res => res.json())
          .then(data => {
            if (!data.success && data.error && data.error.includes('No speaker notes window')) {
              console.log('[Web UI] Speaker notes not open, opening them first for previews...');
              return apiCall('/api/open-speaker-notes').then(() => {
                setTimeout(() => {
                  previewsVisible = true;
                  btn.classList.add('active');
                  container.classList.add('visible');
                  controls.classList.add('with-panel');
                  // If notes are already open, keep compact layout class too
                  if (notesVisible) controls.classList.add('with-notes');
                  body.classList.add('previews-visible');
                  loadSlidePreviews();
                  if (window.previewsRefreshInterval) clearInterval(window.previewsRefreshInterval);
                  window.previewsRefreshInterval = setInterval(loadSlidePreviews, 2000);
                }, 1000);
              });
            }

            // Notes window exists, show previews UI
            previewsVisible = true;
            btn.classList.add('active');
            container.classList.add('visible');
            controls.classList.add('with-panel');
            if (notesVisible) controls.classList.add('with-notes');
            body.classList.add('previews-visible');
            loadSlidePreviews();
            if (window.previewsRefreshInterval) clearInterval(window.previewsRefreshInterval);
            window.previewsRefreshInterval = setInterval(loadSlidePreviews, 2000);
          })
          .catch(err => {
            console.error('[Web UI] Error checking/opening slide previews:', err);
            apiCall('/api/open-speaker-notes').then(() => {
              setTimeout(() => {
                previewsVisible = true;
                btn.classList.add('active');
                container.classList.add('visible');
                controls.classList.add('with-panel');
                if (notesVisible) controls.classList.add('with-notes');
                body.classList.add('previews-visible');
                loadSlidePreviews();
                if (window.previewsRefreshInterval) clearInterval(window.previewsRefreshInterval);
                window.previewsRefreshInterval = setInterval(loadSlidePreviews, 2000);
              }, 1000);
            });
          });
      } else {
        closePreviewsUi();
      }
    });
    
    // Notes zoom controls
    // Allow continuous zoom - the API zooms the actual notes window, and we update Web UI display
    document.getElementById('notes-zoom-out').addEventListener('click', () => {
      // Decrease zoom level (minimum 0.5x)
      if (notesZoomLevel > 0.5) {
        notesZoomLevel = Math.max(0.5, notesZoomLevel - 0.5);
      }
      updateNotesZoom();
      apiCall('/api/zoom-out-notes').catch(() => {}); // Zoom the actual notes window
    });
    
    document.getElementById('notes-zoom-in').addEventListener('click', () => {
      // Increase zoom level (no maximum limit)
      notesZoomLevel += 0.5;
      updateNotesZoom();
      apiCall('/api/zoom-in-notes').catch(() => {}); // Zoom the actual notes window
    });
    
    document.getElementById('remote-btn-next').addEventListener('click', () => {
      triggerHapticFeedback();
      apiCall('/api/next-slide').then(() => {
        // Refresh notes after slide change
        if (notesVisible) {
          loadSpeakerNotes();
        }
        // Refresh previews after slide change
        if (previewsVisible) {
          loadSlidePreviews();
        }
        // Update slide button text
        updateSlideButtons();
      });
    });
    
    document.getElementById('remote-btn-prev').addEventListener('click', () => {
      triggerHapticFeedback();
      apiCall('/api/previous-slide').then(() => {
        // Refresh notes after slide change
        if (notesVisible) {
          loadSpeakerNotes();
        }
        // Refresh previews after slide change
        if (previewsVisible) {
          loadSlidePreviews();
        }
        // Update slide button text
        updateSlideButtons();
      });
    });
    
    // Clickable preview images: current → previous slide, next → next slide
    function refreshAfterSlideChange() {
      if (notesVisible) loadSpeakerNotes();
      if (previewsVisible) loadSlidePreviews();
      updateSlideButtons();
    }
    const currentCard = document.getElementById('slide-preview-current-card');
    const nextCard = document.getElementById('slide-preview-next-card');
    if (currentCard) {
      currentCard.addEventListener('click', () => {
        triggerHapticFeedback();
        apiCall('/api/previous-slide').then(refreshAfterSlideChange);
      });
    }
    if (nextCard) {
      nextCard.addEventListener('click', () => {
        triggerHapticFeedback();
        apiCall('/api/next-slide').then(refreshAfterSlideChange);
      });
    }
    
    // Function to update slide button text with current slide information
    function updateSlideButtons() {
      fetch(API_BASE + '/api/status')
        .then(res => {
          if (!res.ok) {
            throw new Error('HTTP error! status: ' + res.status);
          }
          return res.json();
        })
        .then(data => {
          const prevBtn = document.getElementById('remote-btn-prev');
          const nextBtn = document.getElementById('remote-btn-next');
          const prevBtnControls = document.getElementById('btn-prev-slide');
          const nextBtnControls = document.getElementById('btn-next-slide');
          
          // SVG icons for buttons
          const prevIcon = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="15 18 9 12 15 6"></polyline></svg>';
          const nextIcon = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="9 18 15 12 9 6"></polyline></svg>';
          const prevIconSmall = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"></polyline></svg>';
          const nextIconSmall = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"></polyline></svg>';
          
          updateRemoteSlideContext(data.currentSlide, data.totalSlides);
          
          // Update Remote tab (V2-C: fixed labels; slide position is in the header counter)
          if (prevBtn) {
            prevBtn.innerHTML = prevIcon + '<span class="remote-btn-label">Previous</span>';
          }
          if (nextBtn) {
            nextBtn.innerHTML = '<span class="remote-btn-label">Next slide</span>' + nextIcon;
          }
          
          // Update Controls tab Previous button
          if (prevBtnControls) {
            if (data.previousSlide) {
              prevBtnControls.innerHTML = prevIconSmall + ' Previous Slide (' + data.previousSlide + ')';
            } else {
              prevBtnControls.innerHTML = prevIconSmall + ' Previous Slide';
            }
          }
          
          // Update Controls tab Next button
          if (nextBtnControls) {
            if (data.nextSlide) {
              nextBtnControls.innerHTML = nextIconSmall + ' Next Slide (' + data.nextSlide + ')';
            } else {
              nextBtnControls.innerHTML = nextIconSmall + ' Next Slide';
            }
          }

          // Gate preview button on notes layout
          const previewBtn = document.getElementById('previews-toggle-btn');
          if (previewBtn) {
            const previewsBlocked = (data.notesLayout || 'hide') === 'hide';
            previewBtn.disabled = previewsBlocked;
            if (previewsBlocked) {
              previewBtn.title = 'Change layout to Google Default (50/50) to use previews';
              previewBtn.classList.add('btn-disabled');
            } else {
              previewBtn.title = 'Toggle slide previews';
              previewBtn.classList.remove('btn-disabled');
            }
          }
        })
        .catch(err => {
          // Silently fail - connection might be down, don't spam logs
          console.debug('[Web UI] Failed to update slide buttons:', err.message);
        });
    }
    
    // Helper function to validate and open presentation
    function openPresentation(url, withNotes = false) {
      if (!url) {
        showStatus('Please enter a Google Slides URL', true);
        document.getElementById('presentation-url').focus();
        return;
      }
      
      // Validate it looks like a Google Slides URL
      if (!url.includes('docs.google.com/presentation')) {
        showStatus('Please enter a valid Google Slides URL', true);
        document.getElementById('presentation-url').focus();
        return;
      }
      
      const endpoint = withNotes ? '/api/open-presentation-with-notes' : '/api/open-presentation';
      const btnId = withNotes ? 'btn-open-presentation-with-notes' : 'btn-open-presentation';
      const btn = document.getElementById(btnId);
      const originalText = btn.innerHTML;
      
      // Disable both buttons during request
      document.getElementById('btn-open-presentation').disabled = true;
      document.getElementById('btn-open-presentation-with-notes').disabled = true;
      btn.innerHTML = 'Opening...';
      
      fetch(API_BASE + endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url: url })
      })
        .then(res => {
          if (!res.ok) {
            return res.json().then(data => {
              throw new Error(data.error || 'HTTP error! status: ' + res.status);
            });
          }
          return res.json();
        })
        .then(result => {
          if (result.success) {
            showStatus(result.message || 'Presentation opened successfully!', false);
            document.getElementById('presentation-url').value = ''; // Clear the input
          } else {
            showStatus('Failed to open: ' + (result.error || 'Unknown error'), true);
          }
        })
        .catch(err => {
          console.error('Open presentation error:', err);
          showStatus('Failed to open presentation: ' + err.message + ' (Make sure the app is running)', true);
        })
        .finally(() => {
          document.getElementById('btn-open-presentation').disabled = false;
          document.getElementById('btn-open-presentation-with-notes').disabled = false;
          document.getElementById('btn-open-presentation').innerHTML = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;"><polyline points="5 12 3 12 12 3 21 12 19 12"></polyline><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path><polyline points="9 21 9 12 15 12 15 21"></polyline></svg>Launch Presentation';
          document.getElementById('btn-open-presentation-with-notes').innerHTML = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;"><polyline points="5 12 3 12 12 3 21 12 19 12"></polyline><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path><polyline points="9 21 9 12 15 12 15 21"></polyline></svg>Launch with Notes';
        });
    }
    
    // Open presentation button (without notes)
    document.getElementById('btn-open-presentation').addEventListener('click', () => {
      const url = document.getElementById('presentation-url').value.trim();
      openPresentation(url, false);
    });
    
    // Open presentation with notes button
    document.getElementById('btn-open-presentation-with-notes').addEventListener('click', () => {
      const url = document.getElementById('presentation-url').value.trim();
      openPresentation(url, true);
    });
    
    // Start notes button
    document.getElementById('btn-start-notes').addEventListener('click', () => {
      apiCall('/api/open-speaker-notes');
    });
    
    // Scroll notes (scrolls on presentation machine only)
    document.getElementById('btn-scroll-notes-up').addEventListener('click', () => {
      apiCall('/api/scroll-notes-up');
    });
    document.getElementById('btn-scroll-notes-down').addEventListener('click', () => {
      apiCall('/api/scroll-notes-down');
    });
    
    // Allow Enter key to trigger open (without notes)
    document.getElementById('presentation-url').addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        document.getElementById('btn-open-presentation').click();
      }
    });
    
    // Speaker notes controls removed from default Controls tab - moved to Settings if needed later
    
    // Function to create preset buttons (uses presetUrls array)
    function createPresetButtons(data) {
      const container = document.getElementById('preset-buttons-container');
      container.innerHTML = '';
      const urls = Array.isArray(data?.presetUrls) ? data.presetUrls : [];
      urls.forEach((presetUrl, idx) => {
        if (!presetUrl || presetUrl.trim() === '') return;
        const i = idx + 1;
        const presetGroup = document.createElement('div');
        presetGroup.className = 'web-preset-launch-row';
        const label = document.createElement('div');
        label.className = 'web-preset-launch-label';
        label.textContent = 'Presentation ' + i + ':';
        const buttonGroup = document.createElement('div');
        buttonGroup.className = 'web-preset-launch-actions';
        const launchBtn = document.createElement('button');
        launchBtn.type = 'button';
        launchBtn.className = 'btn';
        launchBtn.style.cssText = 'flex: 1;';
        launchBtn.innerHTML = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;"><polyline points="5 12 3 12 12 3 21 12 19 12"></polyline><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path><polyline points="9 21 9 12 15 12 15 21"></polyline></svg>Launch';
        launchBtn.addEventListener('click', () => { openPresentation(presetUrl, false); });
        const launchWithNotesBtn = document.createElement('button');
        launchWithNotesBtn.type = 'button';
        launchWithNotesBtn.className = 'btn';
        launchWithNotesBtn.style.cssText = 'flex: 1;';
        launchWithNotesBtn.innerHTML = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;"><polyline points="5 12 3 12 12 3 21 12 19 12"></polyline><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path><polyline points="9 21 9 12 15 12 15 21"></polyline></svg>Launch with Notes';
        launchWithNotesBtn.addEventListener('click', () => { openPresentation(presetUrl, true); });
        buttonGroup.appendChild(launchBtn);
        buttonGroup.appendChild(launchWithNotesBtn);
        presetGroup.appendChild(label);
        presetGroup.appendChild(buttonGroup);
        container.appendChild(presetGroup);
      });
      if (container.children.length === 0) {
        if (window.__GSO_WEB_UI_RESTRICTED__) {
          container.innerHTML = '<div class="web-preset-empty">No preset presentations configured.</div>';
        } else {
          container.innerHTML = '<div class="web-preset-empty">No preset presentations configured. <button type="button" class="web-preset-empty-link">Add presets in Settings</button></div>';
        }
      }
    }
    
    function webAddPresetRow(initialValue) {
      const list = document.getElementById('web-preset-list');
      if (!list) return;
      const row = document.createElement('div');
      row.style.display = 'flex';
      row.style.gap = '10px';
      row.style.alignItems = 'center';
      const input = document.createElement('input');
      input.type = 'text';
      input.placeholder = 'https://docs.google.com/presentation/d/...';
      input.value = initialValue || '';
      input.style.flex = '1';
      input.setAttribute('data-web-preset-url', 'true');
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'btn btn-secondary';
      removeBtn.textContent = 'Remove';
      removeBtn.style.padding = '8px 10px';
      removeBtn.addEventListener('click', function() {
        const rows = list.querySelectorAll('[data-web-preset-row]');
        if (rows.length <= 1) { input.value = ''; return; }
        row.remove();
      });
      row.setAttribute('data-web-preset-row', 'true');
      row.appendChild(input);
      row.appendChild(removeBtn);
      list.appendChild(row);
    }
    function webRenderPresetList(urls) {
      const list = document.getElementById('web-preset-list');
      if (!list) return;
      list.innerHTML = '';
      const arr = Array.isArray(urls) ? urls : [];
      if (arr.length === 0) { webAddPresetRow(''); return; }
      arr.forEach(u => webAddPresetRow(u));
    }
    function webGetPresetUrls() {
      const list = document.getElementById('web-preset-list');
      if (!list) return [];
      const inputs = list.querySelectorAll('input[data-web-preset-url="true"]');
      return Array.from(inputs).map(inp => (inp.value || '').trim()).filter(Boolean);
    }
    
    // Test API connection on page load
    fetch(API_BASE + '/api/status')
      .then(res => {
        if (!res.ok) {
          throw new Error('HTTP error! status: ' + res.status);
        }
        return res.json();
      })
      .then(data => {
        console.log('[Web UI] API connection successful:', data);
        // Update slide buttons on initial load
        updateSlideButtons();
        updateNotesZoom();
        // API is reachable, now load presets
        return fetch(API_BASE + '/api/presets');
      })
      .then(res => {
        if (!res.ok) {
          throw new Error('HTTP error! status: ' + res.status);
        }
        return res.json();
      })
      .then(data => {
        webRenderPresetList(data.presetUrls || []);
        createPresetButtons(data);
      })
      .catch(err => {
        console.error('[Web UI] Failed to connect to API:', err);
        console.error('[Web UI] API_BASE was:', API_BASE);
        // Show a warning if API is not reachable
        const statusEl = document.getElementById('status');
        if (statusEl) {
          statusEl.textContent = 'Warning: Cannot connect to API server at ' + API_BASE + '. Controls may not work.';
          statusEl.className = 'status error';
          setTimeout(() => {
            statusEl.className = 'status';
            statusEl.textContent = '';
          }, 10000);
        }
      });
    
    // Poll for slide updates every 2 seconds
    let slideUpdateInterval = setInterval(updateSlideButtons, 2000);
    
    // Clear interval when page unloads
    window.addEventListener('beforeunload', () => {
      if (slideUpdateInterval) {
        clearInterval(slideUpdateInterval);
      }
      if (window.notesRefreshInterval) {
        clearInterval(window.notesRefreshInterval);
        window.notesRefreshInterval = null;
      }
      if (window.previewsRefreshInterval) {
        clearInterval(window.previewsRefreshInterval);
        window.previewsRefreshInterval = null;
      }
    });
    
    const webAddPresetBtn = document.getElementById('web-add-preset');
    if (webAddPresetBtn) {
      webAddPresetBtn.addEventListener('click', () => { webAddPresetRow(''); });
    }
    if (loadBtn) loadBtn.addEventListener('click', () => {
      fetch(API_BASE + '/api/presets')
        .then(res => {
          if (!res.ok) throw new Error('HTTP error! status: ' + res.status);
          return res.json();
        })
        .then(data => {
          webRenderPresetList(data.presetUrls || []);
          showStatus('Presets loaded', false);
          createPresetButtons(data);
        })
        .catch(err => {
          console.error('Load error:', err);
          showStatus('Failed to load presets: ' + err.message + ' (Make sure the app is running)', true);
        });
    });
    
    // Stagetimer integration - Socket.io based
    let stagetimerSocket = null;
    let stagetimerDisplayInterval = null; // For local time updates
    let stagetimerEnabled = false;
    let stagetimerVisible = true;
    let stagetimerState = null; // Store timer state for local calculation
    let stagetimerMessages = []; // Store messages separately
    let stagetimerCurrentTimer = null; // Current timer info (name, speaker, etc.)
    let stagetimerRoomIdCached = '';
    let stagetimerApiKeyCached = '';
    
    function loadStagetimerSettings() {
      fetch(API_BASE + '/api/stagetimer-settings')
        .then(res => res.json())
        .then(data => {
          stagetimerRoomIdCached = String(data.roomId || '').trim();
          stagetimerApiKeyCached = String(data.apiKey || '').trim();
          const roomEl = document.getElementById('stagetimer-room-id');
          const keyEl = document.getElementById('stagetimer-api-key');
          const enabledEl = document.getElementById('stagetimer-enabled');
          const visibleEl = document.getElementById('stagetimer-visible');
          if (roomEl) roomEl.value = stagetimerRoomIdCached;
          if (keyEl) keyEl.value = stagetimerApiKeyCached;
          if (enabledEl) enabledEl.checked = data.enabled !== false;
          if (visibleEl) visibleEl.checked = data.visible !== false;
          stagetimerEnabled = data.enabled !== false;
          stagetimerVisible = data.visible !== false;
          
          // Update display based on visibility and configuration
          updateStagetimerVisibility();
          
          if (stagetimerEnabled && stagetimerRoomIdCached && stagetimerApiKeyCached) {
            connectStagetimerSocket(stagetimerRoomIdCached, stagetimerApiKeyCached);
          } else {
            disconnectStagetimerSocket();
            updateStagetimerDisplay(null, 'Not configured');
          }
        })
        .catch(err => {
          console.error('Failed to load stagetimer settings:', err);
        });
    }
    
    function updateStagetimerVisibility() {
      const container = document.getElementById('stagetimer-container');
      if (!container) return;
      const keyEl = document.getElementById('stagetimer-api-key');
      const roomEl = document.getElementById('stagetimer-room-id');
      const hasApiKey = keyEl ? keyEl.value.trim().length > 0 : stagetimerApiKeyCached.length > 0;
      const hasRoomId = roomEl ? roomEl.value.trim().length > 0 : stagetimerRoomIdCached.length > 0;
      
      // Hide if: not visible OR not enabled OR missing API key/room ID
      if (!stagetimerVisible || !stagetimerEnabled || !hasApiKey || !hasRoomId) {
        container.style.display = 'none';
      } else {
        container.style.display = 'block';
      }
    }
    
    function saveStagetimerSettings() {
      const roomId = document.getElementById('stagetimer-room-id').value.trim();
      const apiKey = document.getElementById('stagetimer-api-key').value.trim();
      const enabled = document.getElementById('stagetimer-enabled').checked;
      const visible = document.getElementById('stagetimer-visible').checked;
      
      fetch(API_BASE + '/api/stagetimer-settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ roomId, apiKey, enabled, visible })
      })
        .then(res => res.json())
        .then(result => {
          if (result.success) {
            showStatus('Stagetimer settings saved', false);
            stagetimerEnabled = enabled;
            stagetimerVisible = visible;
            
            // Update visibility
            updateStagetimerVisibility();
            
            if (stagetimerEnabled && roomId && apiKey) {
              connectStagetimerSocket(roomId, apiKey);
            } else {
              disconnectStagetimerSocket();
              updateStagetimerDisplay(null, enabled ? 'Please configure Room ID and API Key' : 'Disabled');
            }
          } else {
            showStatus('Failed to save: ' + (result.error || 'Unknown error'), true);
          }
        })
        .catch(err => {
          console.error('Save stagetimer settings error:', err);
          showStatus('Failed to save settings: ' + err.message, true);
        });
    }
    
    function updateStagetimerDisplay(data, errorMessage) {
      const container = document.getElementById('stagetimer-container');
      const labelEl = document.getElementById('stagetimer-label');
      const timeEl = document.getElementById('stagetimer-time');
      const statusEl = document.getElementById('stagetimer-status');
      const messagesEl = document.getElementById('stagetimer-messages');
      
      // Ensure we never "perma-hide" messages via inline styles
      if (messagesEl && messagesEl.style) {
        messagesEl.style.removeProperty('display');
      }
      
      // Check visibility and configuration
      updateStagetimerVisibility();
      
      // If not visible or not enabled, don't update content
      if (!stagetimerVisible || !stagetimerEnabled) {
        return;
      }
      
      // If there's an error and no API key, hide it
      if (errorMessage && errorMessage.includes('not configured')) {
        container.style.display = 'none';
        return;
      }
      
      container.style.display = 'block';
      
      if (errorMessage || !data || !data.success) {
        container.className = 'stagetimer-container error';
        labelEl.textContent = data?.timerName || ''; // Still try to show timer name if available
        timeEl.textContent = '--:--';
        statusEl.textContent = errorMessage || 'Error loading timer';
        if (messagesEl) {
          messagesEl.innerHTML = '';
          messagesEl.classList.remove('visible');
        }
        return;
      }
      
      // Update timer name (label) - use name from data or current timer, fallback to empty
      labelEl.textContent = data.timerName || stagetimerCurrentTimer?.name || '';
      
      timeEl.textContent = data.displayTime || '0:00';
      
      // Show speaker name below the timer
      statusEl.textContent = data.speaker || stagetimerCurrentTimer?.speaker || '';
      
      // Determine state and styling (but don't show status text)
      if (data.remainingMs !== undefined) {
        const remainingSeconds = Math.floor(data.remainingMs / 1000);
        if (remainingSeconds < 0) {
          container.className = 'stagetimer-container overtime';
        } else if (remainingSeconds <= 15) {
          container.className = 'stagetimer-container critical';
        } else if (remainingSeconds <= 60) {
          container.className = 'stagetimer-container warning';
        } else if (data.running) {
          container.className = 'stagetimer-container running';
        } else {
          container.className = 'stagetimer-container';
        }
      } else if (data.running) {
        container.className = 'stagetimer-container running';
      } else if (data.pause) {
        container.className = 'stagetimer-container';
      } else {
        container.className = 'stagetimer-container';
      }
      
      // Display messages - positioned absolutely so buttons don't move
      console.log('[Stagetimer Display] Updating display, messages:', data.messages);
      if (data.messages && data.messages.length > 0) {
        console.log('[Stagetimer Display] Showing', data.messages.length, 'messages');
        if (!messagesEl) return;
        messagesEl.innerHTML = '';
        data.messages.forEach((msg, index) => {
          console.log('[Stagetimer Display] Message', index + ':', msg);
          const messageDiv = document.createElement('div');
          messageDiv.className = 'stagetimer-message ' + (msg.color || 'white');
          if (msg.bold) messageDiv.classList.add('bold');
          if (msg.uppercase) messageDiv.classList.add('uppercase');
          messageDiv.textContent = msg.text || '';
          messagesEl.appendChild(messageDiv);
        });
        messagesEl.classList.add('visible');
      } else {
        console.log('[Stagetimer Display] No messages to display');
        if (messagesEl) {
          messagesEl.innerHTML = '';
          messagesEl.classList.remove('visible');
        }
      }
    }
    
    // Socket.io connection for real-time stagetimer updates
    function connectStagetimerSocket(roomId, apiKey) {
      // Disconnect existing connection if any
      disconnectStagetimerSocket();
      
      if (!window.io) {
        console.error('[Stagetimer] Socket.io library not loaded');
        updateStagetimerDisplay(null, 'Socket.io library not available');
        return;
      }
      
      console.log('[Stagetimer] Connecting to socket.io...');
      
      try {
        stagetimerSocket = io('https://api.stagetimer.io', {
          path: '/v1/socket.io',
          auth: {
            room_id: roomId,
            api_key: apiKey
          },
          transports: ['websocket', 'polling'], // Try websocket first, fallback to polling
          reconnection: true,
          reconnectionDelay: 1000,
          reconnectionDelayMax: 5000,
          reconnectionAttempts: Infinity
        });
        
        // Connection successful
        stagetimerSocket.on('connect', () => {
          console.log('[Stagetimer] Socket.io connected');
          updateStagetimerDisplay(null, null); // Clear any error messages
          
          // Start local time updates
          if (!stagetimerDisplayInterval) {
            stagetimerDisplayInterval = setInterval(updateStagetimerDisplayFromState, 1000);
          }
        });
        
        // Connection error
        stagetimerSocket.on('connect_error', (error) => {
          console.error('[Stagetimer] Socket.io connection error:', error);
          updateStagetimerDisplay(null, 'Connection error: ' + (error.message || 'Failed to connect'));
        });
        
        // Disconnected
        stagetimerSocket.on('disconnect', (reason) => {
          console.warn('[Stagetimer] Socket.io disconnected:', reason);
          if (reason === 'io server disconnect') {
            // Server disconnected, try to reconnect
            stagetimerSocket.connect();
          }
        });
        
        // Reconnection attempt
        stagetimerSocket.on('reconnect_attempt', (attemptNumber) => {
          console.log('[Stagetimer] Reconnection attempt', attemptNumber);
        });
        
        // Reconnected
        stagetimerSocket.on('reconnect', (attemptNumber) => {
          console.log('[Stagetimer] Reconnected after', attemptNumber, 'attempts');
        });
        
        // Playback status updates (timer start/stop/pause/reset)
        stagetimerSocket.on('playback_status', (data) => {
          console.log('[Stagetimer] playback_status event:', data);
          
          if (data && data._model === 'playback_status') {
            // Update state with new playback status
            const now = data.server_time || (data._updated_at ? new Date(data._updated_at).getTime() : Date.now());
            
            // Preserve existing timer info if state already exists
            const existingTimerName = stagetimerState?.timerName || stagetimerCurrentTimer?.name || '';
            const existingSpeaker = stagetimerState?.speaker || stagetimerCurrentTimer?.speaker || '';
            const existingMessages = stagetimerState?.messages || stagetimerMessages.filter(m => m.showing).map(m => ({
              text: m.text || '',
              color: m.color || 'white',
              bold: m.bold || false,
              uppercase: m.uppercase || false
            })) || [];
            
            stagetimerState = {
              success: true,
              running: data.running || false,
              start: data.start,
              finish: data.finish,
              pause: data.pause,
              serverTime: now,
              timerId: data.timer_id,
              timerName: existingTimerName,
              speaker: existingSpeaker,
              messages: existingMessages,
              lastSyncTime: Date.now()
            };
            
            // Update display immediately
            updateStagetimerDisplayFromState();
          }
        });
        
        // Current timer updates (name, speaker, notes, etc.)
        stagetimerSocket.on('current_timer', (data) => {
          console.log('[Stagetimer] current_timer event:', data);
          
          if (data && data._model === 'timer') {
            stagetimerCurrentTimer = {
              timerId: data._id,
              name: data.name || '',
              speaker: data.speaker || '',
              notes: data.notes || ''
            };
            
            // Update state with timer info
            if (stagetimerState) {
              stagetimerState.timerName = stagetimerCurrentTimer.name;
              stagetimerState.speaker = stagetimerCurrentTimer.speaker;
            } else {
              // If no state yet, create a minimal state (will be updated by playback_status)
              stagetimerState = {
                success: true,
                running: false,
                timerName: stagetimerCurrentTimer.name,
                speaker: stagetimerCurrentTimer.speaker,
                messages: stagetimerMessages.filter(m => m.showing).map(m => ({
                  text: m.text || '',
                  color: m.color || 'white',
                  bold: m.bold || false,
                  uppercase: m.uppercase || false
                })) || [],
                lastSyncTime: Date.now()
              };
            }
            
            // Update display
            updateStagetimerDisplayFromState();
          }
        });
        
        // Message updates (show/hide/update)
        stagetimerSocket.on('message', (data) => {
          console.log('[Stagetimer] message event:', data);
          
          if (data && data._model === 'message') {
            // Update messages array
            if (data.showing) {
              // Add or update message
              const existingIndex = stagetimerMessages.findIndex(m => m._id === data._id);
              if (existingIndex >= 0) {
                stagetimerMessages[existingIndex] = data;
              } else {
                stagetimerMessages.push(data);
              }
            } else {
              // Remove message
              stagetimerMessages = stagetimerMessages.filter(m => m._id !== data._id);
            }
            
            // Update state
            if (stagetimerState) {
              stagetimerState.messages = stagetimerMessages.filter(m => m.showing).map(m => ({
                text: m.text || '',
                color: m.color || 'white',
                bold: m.bold || false,
                uppercase: m.uppercase || false
              }));
            } else {
              // If no state yet, create a minimal state
              stagetimerState = {
                success: true,
                running: false,
                timerName: stagetimerCurrentTimer?.name || '',
                speaker: stagetimerCurrentTimer?.speaker || '',
                messages: stagetimerMessages.filter(m => m.showing).map(m => ({
                  text: m.text || '',
                  color: m.color || 'white',
                  bold: m.bold || false,
                  uppercase: m.uppercase || false
                })) || [],
                lastSyncTime: Date.now()
              };
            }
            
            // Update display
            updateStagetimerDisplayFromState();
          }
        });
        
        // Room updates (blackout, focus, on-air, etc.)
        stagetimerSocket.on('room', (data) => {
          console.log('[Stagetimer] room event:', data);
          // We don't currently use room state, but log it for debugging
        });
        
        // Flash events
        stagetimerSocket.on('flash', (data) => {
          console.log('[Stagetimer] flash event:', data);
          // We don't currently handle flash events, but log them
        });
        
        // Listen to all events for debugging
        stagetimerSocket.onAny((event, ...args) => {
          console.log('[Stagetimer] Event received:', event, args);
        });
        
      } catch (error) {
        console.error('[Stagetimer] Error creating socket connection:', error);
        updateStagetimerDisplay(null, 'Failed to connect: ' + error.message);
      }
    }
    
    function disconnectStagetimerSocket() {
      if (stagetimerSocket) {
        console.log('[Stagetimer] Disconnecting socket...');
        stagetimerSocket.disconnect();
        stagetimerSocket = null;
      }
      
      if (stagetimerDisplayInterval) {
        clearInterval(stagetimerDisplayInterval);
        stagetimerDisplayInterval = null;
      }
      
      // Clear state
      stagetimerState = null;
      stagetimerMessages = [];
      stagetimerCurrentTimer = null;
    }
    
    // Calculate and display time locally based on stored state
    function updateStagetimerDisplayFromState() {
      // Check visibility and configuration
      if (!stagetimerVisible || !stagetimerEnabled) {
        return;
      }
      
      if (!stagetimerState || !stagetimerState.success) {
        return; // No state to work with
      }
      
      const state = stagetimerState;
      const now = Date.now();
      
      // Calculate time difference since last server sync
      const timeSinceSync = now - (state.lastSyncTime || now);
      
      // Calculate remaining/elapsed time
      let remainingMs = 0;
      let elapsedMs = 0;
      let displayTime = '0:00';
      let isRunning = state.running || false;
      
      if (state.finish && state.start) {
        const duration = state.finish - state.start;
        
        if (isRunning) {
          // Timer is running - calculate based on server time + elapsed local time
          const serverTimeAtSync = state.serverTime || state.lastSyncTime;
          const localTimeAtSync = state.lastSyncTime;
          const adjustedNow = serverTimeAtSync + (now - localTimeAtSync);
          remainingMs = state.finish - adjustedNow; // Allow negative values
          elapsedMs = adjustedNow - state.start;
        } else if (state.pause) {
          // Timer is paused - use stored values
          elapsedMs = state.pause - state.start;
          remainingMs = duration - elapsedMs;
        } else {
          // Timer not started
          remainingMs = duration;
          elapsedMs = 0;
        }
        
        // Format time as MM:SS or HH:MM:SS (allow negative)
        const totalSeconds = Math.floor(remainingMs / 1000);
        const isNegative = totalSeconds < 0;
        const absSeconds = Math.abs(totalSeconds);
        const hours = Math.floor(absSeconds / 3600);
        const minutes = Math.floor((absSeconds % 3600) / 60);
        const seconds = absSeconds % 60;
        
        const sign = isNegative ? '-' : '';
        const minStr = String(minutes).padStart(2, '0');
        const secStr = String(seconds).padStart(2, '0');
        
        if (hours > 0) {
          displayTime = sign + hours + ':' + minStr + ':' + secStr;
        } else {
          displayTime = sign + minutes + ':' + secStr;
        }
      }
      
      // Update display with calculated time
      const container = document.getElementById('stagetimer-container');
      const labelEl = document.getElementById('stagetimer-label');
      const timeEl = document.getElementById('stagetimer-time');
      const statusEl = document.getElementById('stagetimer-status');
      
      if (!container || !labelEl || !timeEl || !statusEl) return;
      
      labelEl.textContent = state.timerName || stagetimerCurrentTimer?.name || '';
      timeEl.textContent = displayTime;
      statusEl.textContent = state.speaker || stagetimerCurrentTimer?.speaker || '';
      
      const remainingSeconds = Math.floor(remainingMs / 1000);
      if (state.finish && state.start) {
        if (remainingSeconds < 0) {
          container.className = 'stagetimer-container overtime';
        } else if (remainingSeconds <= 15) {
          container.className = 'stagetimer-container critical';
        } else if (remainingSeconds <= 60) {
          container.className = 'stagetimer-container warning';
        } else if (isRunning) {
          container.className = 'stagetimer-container running';
        } else {
          container.className = 'stagetimer-container';
        }
      } else if (isRunning) {
        container.className = 'stagetimer-container running';
      } else {
        container.className = 'stagetimer-container';
      }
      
      // Update messages from stored state
      const messagesEl = document.getElementById('stagetimer-messages');
      if (messagesEl && state.messages && state.messages.length > 0) {
        messagesEl.innerHTML = '';
        state.messages.forEach((msg) => {
          const messageDiv = document.createElement('div');
          messageDiv.className = 'stagetimer-message ' + (msg.color || 'white');
          if (msg.bold) messageDiv.classList.add('bold');
          if (msg.uppercase) messageDiv.classList.add('uppercase');
          messageDiv.textContent = msg.text || '';
          messagesEl.appendChild(messageDiv);
        });
        messagesEl.classList.add('visible');
      } else if (messagesEl) {
        messagesEl.innerHTML = '';
        messagesEl.classList.remove('visible');
      }
    }
    
    if (${webUiDebugConsoleEnabled ? 'true' : 'false'}) {
      // Debug console functionality (enabled from desktop app)
      const debugConsole = document.getElementById('debug-console');
      if (debugConsole) {
        const originalConsoleLog = console.log;
        const originalConsoleError = console.error;
        const originalConsoleWarn = console.warn;
        
        function addToDebugConsole(message, type = 'log') {
          const timestamp = new Date().toLocaleTimeString();
          const color = type === 'error' ? '#f44336' : type === 'warn' ? '#ff9800' : '#4caf50';
          const prefix = type === 'error' ? '[ERROR]' : type === 'warn' ? '[WARN]' : '[LOG]';
          
          const logEntry = document.createElement('div');
          logEntry.style.marginBottom = '4px';
          logEntry.style.color = color;
          logEntry.innerHTML = '<span style="color: #888;">[' + timestamp + ']</span> ' + prefix + ' ' + message;
          
          const readyMsg = debugConsole.querySelector('div[style*="color: #888"]');
          if (readyMsg && readyMsg.textContent.includes('Console ready')) {
            readyMsg.remove();
          }
          
          debugConsole.appendChild(logEntry);
          debugConsole.scrollTop = debugConsole.scrollHeight;
          
          while (debugConsole.children.length > 100) {
            debugConsole.removeChild(debugConsole.firstChild);
          }
        }
        
        console.log = function(...args) {
          originalConsoleLog.apply(console, arguments);
          addToDebugConsole(args.map(arg => typeof arg === 'object' ? JSON.stringify(arg, null, 2) : String(arg)).join(' '), 'log');
        };
        
        console.error = function(...args) {
          originalConsoleError.apply(console, arguments);
          addToDebugConsole(args.map(arg => typeof arg === 'object' ? JSON.stringify(arg, null, 2) : String(arg)).join(' '), 'error');
        };
        
        console.warn = function(...args) {
          originalConsoleWarn.apply(console, arguments);
          addToDebugConsole(args.map(arg => typeof arg === 'object' ? JSON.stringify(arg, null, 2) : String(arg)).join(' '), 'warn');
        };
        
        const clearBtn = document.getElementById('btn-clear-console');
        if (clearBtn) {
          clearBtn.addEventListener('click', () => {
            debugConsole.innerHTML = '<div style="color: #888;">Console cleared...</div>';
          });
        }
      }
    }
    
    // Load stagetimer settings on page load
    loadStagetimerSettings();
    
    // WAN tunnel status + QR controls
    async function loadWebTunnelStatus() {
      try {
        const res = await fetch(API_BASE + '/api/status');
        const data = await res.json();
        const statusEl = document.getElementById('web-tunnel-status');
        const qrRow = document.getElementById('web-tunnel-qr-row');
        if (data.tunnelEnabled && data.tunnelUrl) {
          statusEl.innerHTML = 'Tunnel active: <strong style="word-break:break-all;">' + data.tunnelUrl + '</strong>';
          statusEl.style.background = 'rgba(0,120,200,0.1)';
          statusEl.style.border = '1px solid rgba(0,120,200,0.3)';
          statusEl.style.borderRadius = '6px';
          statusEl.style.padding = '8px 10px';
          qrRow.style.display = 'flex';
        } else if (data.tunnelEnabled) {
          statusEl.textContent = 'Tunnel enabled — connecting…';
          qrRow.style.display = 'none';
        } else {
          statusEl.textContent = 'Tunnel is off. Enable it from the desktop app Settings → WAN Access.';
          statusEl.style.background = '';
          statusEl.style.border = '';
          statusEl.style.padding = '';
          qrRow.style.display = 'none';
        }
      } catch (e) {
        document.getElementById('web-tunnel-status').textContent = 'Could not load tunnel status.';
      }
    }
    loadWebTunnelStatus();

    if (!window.__GSO_WEB_UI_RESTRICTED__) {
    // Save stagetimer settings button
    document.getElementById('btn-save-stagetimer').addEventListener('click', saveStagetimerSettings);
    document.getElementById('btn-load-stagetimer').addEventListener('click', loadStagetimerSettings);

    document.getElementById('btn-show-tunnel-qr').addEventListener('click', async () => {
      const duration = parseInt(document.getElementById('web-tunnel-qr-duration').value) || 20;
      try {
        const res = await fetch(API_BASE + '/api/show-tunnel-qr', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ duration })
        });
        const data = await res.json();
        showStatus(data.message || (data.error ? 'Error: ' + data.error : 'Done'), !!data.error);
      } catch (e) { showStatus('Request failed', true); }
    });

    document.getElementById('btn-hide-tunnel-qr').addEventListener('click', async () => {
      try {
        await fetch(API_BASE + '/api/hide-tunnel-qr', { method: 'POST' });
        showStatus('QR hidden');
      } catch (e) { showStatus('Request failed', true); }
    });
    
    // Load all settings when Settings tab is opened
    let settingsLoaded = false;
    document.querySelector('[data-tab="settings"]').addEventListener('click', () => {
      if (!settingsLoaded) {
        loadAllSettings();
        settingsLoaded = true;
      }
    });
    
    // Load settings immediately if Settings tab is already active
    if (document.getElementById('tab-settings').classList.contains('active')) {
      loadAllSettings();
      settingsLoaded = true;
    }
    
    // Function to load all settings
    let webBackupStatusByIp = {};
    let webBackupHandlersAttached = false;

    function normalizeWebBackupIps(ips) {
      if (!Array.isArray(ips)) return [];
      const out = [];
      const seen = new Set();
      ips.forEach((raw) => {
        const v = String(raw || '').trim();
        if (!v) return;
        if (seen.has(v)) return;
        seen.add(v);
        out.push(v);
      });
      return out;
    }

    function getWebBackupIpListEl() {
      return document.getElementById('web-backup-ip-list');
    }

    function getWebBackupIpInputs() {
      const list = getWebBackupIpListEl();
      if (!list) return [];
      return Array.from(list.querySelectorAll('input[data-web-backup-ip="true"]'));
    }

    function getWebBackupIpsFromUi() {
      return normalizeWebBackupIps(getWebBackupIpInputs().map((el) => String(el.value || '').trim()));
    }

    function setWebBackupStatusBadge(el, ip) {
      if (!el) return;
      const v = String(ip || '').trim();
      const status = v ? webBackupStatusByIp[v] : null;

      if (!v) {
        el.textContent = '-';
        el.style.background = 'transparent';
        el.style.color = '#888';
        return;
      }
      if (status === 'connected') {
        el.textContent = 'Connected';
        el.style.background = '#4caf50';
        el.style.color = 'white';
        return;
      }
      if (status === 'disconnected') {
        el.textContent = 'Disconnected';
        el.style.background = '#f44336';
        el.style.color = 'white';
        return;
      }
      el.textContent = 'Checking...';
      el.style.background = '#ff9800';
      el.style.color = 'white';
    }

    function refreshWebBackupStatusBadges() {
      const list = getWebBackupIpListEl();
      if (!list) return;
      const rows = Array.from(list.querySelectorAll('[data-web-backup-row="true"]'));
      rows.forEach((row) => {
        const input = row.querySelector('input[data-web-backup-ip="true"]');
        const badge = row.querySelector('span[data-web-backup-status="true"]');
        setWebBackupStatusBadge(badge, input ? input.value : '');
      });
    }

    function addWebBackupIpRow(initialValue = '') {
      const list = getWebBackupIpListEl();
      if (!list) return;

      const row = document.createElement('div');
      row.setAttribute('data-web-backup-row', 'true');
      row.style.display = 'flex';
      row.style.gap = '10px';
      row.style.alignItems = 'center';
      row.style.width = '100%';
      row.style.minWidth = '0';

      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'input-field';
      input.placeholder = '192.168.1.100 or hostname';
      input.value = initialValue || '';
      input.setAttribute('data-web-backup-ip', 'true');
      input.style.flex = '1 1 0%';
      input.style.width = '100%';
      input.style.minWidth = '140px';

      const badge = document.createElement('span');
      badge.setAttribute('data-web-backup-status', 'true');
      badge.style.fontSize = '12px';
      badge.style.padding = '4px 8px';
      badge.style.borderRadius = '4px';
      badge.style.minWidth = '90px';
      badge.style.textAlign = 'center';

      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'btn btn-secondary';
      removeBtn.textContent = 'Remove';
      removeBtn.style.padding = '8px 10px';
      removeBtn.style.minWidth = '88px';

      removeBtn.addEventListener('click', () => {
        const rows = list.querySelectorAll('[data-web-backup-row="true"]');
        if (rows.length <= 1) {
          input.value = '';
          refreshWebBackupStatusBadges();
          return;
        }
        row.remove();
        refreshWebBackupStatusBadges();
      });

      input.addEventListener('change', () => {
        refreshWebBackupStatusBadges();
      });

      row.appendChild(input);
      row.appendChild(badge);
      row.appendChild(removeBtn);
      list.appendChild(row);

      setWebBackupStatusBadge(badge, input.value);
    }

    function renderWebBackupIpList(ips = []) {
      const list = getWebBackupIpListEl();
      if (!list) return;
      list.innerHTML = '';
      const normalized = Array.isArray(ips) ? ips.map(v => String(v || '')) : [];
      if (normalized.length === 0) {
        addWebBackupIpRow('');
        return;
      }
      normalized.forEach((ip) => addWebBackupIpRow(ip));
    }

    function attachWebBackupHandlersOnce() {
      if (webBackupHandlersAttached) return;
      webBackupHandlersAttached = true;
      const addBtn = document.getElementById('web-add-backup-ip');
      if (addBtn) {
        addBtn.addEventListener('click', () => {
          addWebBackupIpRow('');
          const inputs = getWebBackupIpInputs();
          if (inputs.length) inputs[inputs.length - 1].focus();
        });
      }
    }

    async function loadAllSettings() {
      try {
        // Load displays
        const displaysRes = await fetch(API_BASE + '/api/displays');
        const displays = await displaysRes.json();
        
        const presentationSelect = document.getElementById('web-presentation-display');
        const notesSelect = document.getElementById('web-notes-display');
        
        presentationSelect.innerHTML = '';
        notesSelect.innerHTML = '';
        
        displays.forEach(display => {
          const option1 = document.createElement('option');
          option1.value = display.id;
          option1.textContent = display.label + (display.primary ? ' (Primary)' : '');
          presentationSelect.appendChild(option1);
          
          const option2 = document.createElement('option');
          option2.value = display.id;
          option2.textContent = display.label + (display.primary ? ' (Primary)' : '');
          notesSelect.appendChild(option2);
        });
        
        // Load preferences
        const prefsRes = await fetch(API_BASE + '/api/preferences');
        const prefs = await prefsRes.json();
        
        // Set display values
        if (prefs.presentationDisplayId) {
          presentationSelect.value = prefs.presentationDisplayId;
        }
        if (prefs.notesDisplayId) {
          notesSelect.value = prefs.notesDisplayId;
        }
        
        // Set notes layout preference
        const notesLayoutSelect = document.getElementById('web-notes-layout');
        if (notesLayoutSelect && prefs.notesLayout) {
          notesLayoutSelect.value = prefs.notesLayout;
        }
        const webDefaultNotesZoom = document.getElementById('web-default-notes-zoom-steps');
        if (webDefaultNotesZoom) {
          const dz = prefs.defaultNotesZoomSteps;
          webDefaultNotesZoom.value = dz !== undefined && dz !== null ? String(dz) : '0';
        }

        // Set machine name
        document.getElementById('web-machine-name').value = prefs.machineName || '';
        
        // Set primary/backup mode
        const mode = prefs.primaryBackupMode || 'standalone';
        document.getElementById('web-mode-primary').checked = mode === 'primary';
        document.getElementById('web-mode-backup').checked = mode === 'backup';
        document.getElementById('web-mode-standalone').checked = mode === 'standalone';
        
        const backupConfig = document.getElementById('web-backup-config');
        if (mode === 'primary') {
          backupConfig.style.display = 'block';
        } else {
          backupConfig.style.display = 'none';
        }
        
        // Set backup configuration (unlimited). Fallback to legacy fields if present.
        document.getElementById('web-backup-port').value = prefs.backupPort || '9595';
        const legacyIps = [prefs.backupIp1, prefs.backupIp2, prefs.backupIp3].filter(v => v && String(v).trim() !== '');
        const backupIps = Array.isArray(prefs.backupIps) ? prefs.backupIps : legacyIps;
        attachWebBackupHandlersOnce();
        renderWebBackupIpList(backupIps);
        refreshWebBackupStatusBadges();
        
        // Set network ports
        document.getElementById('web-api-port').value = prefs.apiPort || '9595';
        document.getElementById('web-web-ui-port').value = prefs.webUiPort || '80';
        
        // Set logging preferences
        const verboseEl = document.getElementById('web-verbose-logging');
        if (verboseEl) {
          verboseEl.checked = prefs.verboseLogging === true;
        }
        
        // Set up primary/backup mode change handlers
        document.getElementById('web-mode-primary').addEventListener('change', () => {
          if (document.getElementById('web-mode-primary').checked) {
            backupConfig.style.display = 'block';
          }
        });
        document.getElementById('web-mode-backup').addEventListener('change', () => {
          if (document.getElementById('web-mode-backup').checked) {
            backupConfig.style.display = 'none';
          }
        });
        document.getElementById('web-mode-standalone').addEventListener('change', () => {
          if (document.getElementById('web-mode-standalone').checked) {
            backupConfig.style.display = 'none';
          }
        });
        
        // Start backup status polling if in primary mode
        if (mode === 'primary') {
          startWebBackupStatusPolling();
        }
      } catch (error) {
        console.error('Failed to load settings:', error);
        showStatus('Failed to load settings: ' + error.message, true);
      }
    }
    
    // Live-save notes layout preference when dropdown changes
    document.getElementById('web-notes-layout').addEventListener('change', async (e) => {
      try {
        const res = await fetch(API_BASE + '/api/preferences', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ notesLayout: e.target.value })
        });
        const result = await res.json();
        if (result.success) {
          showStatus('Notes layout saved. Use Relaunch Notes to apply.', false);
        } else {
          showStatus('Failed to save layout: ' + (result.error || 'Unknown error'), true);
        }
      } catch (error) {
        showStatus('Failed to save layout: ' + error.message, true);
      }
    });

    // Relaunch Notes button — close + reopen notes window to apply layout change
    document.getElementById('btn-relaunch-notes').addEventListener('click', async () => {
      const btn = document.getElementById('btn-relaunch-notes');
      const originalText = btn.textContent;
      btn.disabled = true;
      btn.textContent = 'Relaunching...';
      try {
        const res = await fetch(API_BASE + '/api/relaunch-speaker-notes', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({})
        });
        const result = await res.json();
        if (result.success) {
          showStatus('Notes relaunched with new layout', false);
        } else {
          showStatus('Relaunch failed: ' + (result.error || 'Unknown error'), true);
        }
      } catch (error) {
        showStatus('Relaunch failed: ' + error.message, true);
      } finally {
        btn.disabled = false;
        btn.textContent = originalText;
      }
    });

    // Save monitor settings
    document.getElementById('btn-save-displays').addEventListener('click', async () => {
      try {
        const webZoomEl = document.getElementById('web-default-notes-zoom-steps');
        let defaultNotesZoomSteps = 0;
        if (webZoomEl) {
          const pz = parseInt(webZoomEl.value, 10);
          defaultNotesZoomSteps = Number.isNaN(pz) ? 0 : pz;
        }
        const prefs = {
          presentationDisplayId: parseInt(document.getElementById('web-presentation-display').value),
          notesDisplayId: parseInt(document.getElementById('web-notes-display').value),
          notesLayout: document.getElementById('web-notes-layout').value,
          defaultNotesZoomSteps
        };
        
        const res = await fetch(API_BASE + '/api/preferences', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(prefs)
        });
        
        const result = await res.json();
        if (result.success) {
          showStatus('Monitor settings saved', false);
        } else {
          showStatus('Failed to save monitor settings: ' + (result.error || 'Unknown error'), true);
        }
      } catch (error) {
        showStatus('Failed to save monitor settings: ' + error.message, true);
      }
    });
    
    // Save machine name
    document.getElementById('btn-save-machine-name').addEventListener('click', async () => {
      try {
        const prefs = {
          machineName: document.getElementById('web-machine-name').value.trim()
        };
        
        const res = await fetch(API_BASE + '/api/preferences', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(prefs)
        });
        
        const result = await res.json();
        if (result.success) {
          showStatus('Machine name saved', false);
          // Reload page to update header
          setTimeout(() => {
            window.location.reload();
          }, 1000);
        } else {
          showStatus('Failed to save machine name: ' + (result.error || 'Unknown error'), true);
        }
      } catch (error) {
        showStatus('Failed to save machine name: ' + error.message, true);
      }
    });
    
    // Save primary/backup settings
    document.getElementById('btn-save-primary-backup').addEventListener('click', async () => {
      try {
        let mode = 'standalone';
        if (document.getElementById('web-mode-primary').checked) {
          mode = 'primary';
        } else if (document.getElementById('web-mode-backup').checked) {
          mode = 'backup';
        }
        
        const backupPort = parseInt(document.getElementById('web-backup-port').value);
        if (mode === 'primary' && (isNaN(backupPort) || backupPort < 1024 || backupPort > 65535)) {
          showStatus('Backup port must be between 1024 and 65535', true);
          return;
        }
        
        const prefs = { primaryBackupMode: mode };
        if (mode === 'primary') {
          prefs.backupPort = backupPort;
          prefs.backupIps = getWebBackupIpsFromUi();
        }
        
        const res = await fetch(API_BASE + '/api/preferences', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(prefs)
        });
        
        const result = await res.json();
        if (result.success) {
          showStatus('Primary/Backup settings saved', false);
          
          // Restart backup status polling if needed
          if (mode === 'primary') {
            startWebBackupStatusPolling();
          } else {
            stopWebBackupStatusPolling();
          }
        } else {
          showStatus('Failed to save Primary/Backup settings: ' + (result.error || 'Unknown error'), true);
        }
      } catch (error) {
        showStatus('Failed to save Primary/Backup settings: ' + error.message, true);
      }
    });
    
    // Save port settings
    document.getElementById('btn-save-ports').addEventListener('click', async () => {
      try {
        const apiPort = parseInt(document.getElementById('web-api-port').value);
        const webUiPort = parseInt(document.getElementById('web-web-ui-port').value);
        
        if (isNaN(apiPort) || apiPort < 1024 || apiPort > 65535) {
          showStatus('API port must be between 1024 and 65535', true);
          return;
        }
        
        if (isNaN(webUiPort) || webUiPort < 1 || webUiPort > 65535) {
          showStatus('Web UI port must be between 1 and 65535', true);
          return;
        }
        
        const prefs = {
          apiPort: apiPort,
          webUiPort: webUiPort
        };
        if (document.getElementById('web-mode-primary').checked) {
          prefs.backupPort = apiPort;
          const wbp = document.getElementById('web-backup-port');
          if (wbp) wbp.value = String(apiPort);
        }

        const res = await fetch(API_BASE + '/api/preferences', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(prefs)
        });
        
        const result = await res.json();
        if (result.success) {
          showStatus('Port settings saved. Please restart the app for changes to take effect.', false);
        } else {
          showStatus('Failed to save port settings: ' + (result.error || 'Unknown error'), true);
        }
      } catch (error) {
        showStatus('Failed to save port settings: ' + error.message, true);
      }
    });
    
    // Save logging settings
    const saveLoggingBtn = document.getElementById('btn-save-logging');
    if (saveLoggingBtn) {
      saveLoggingBtn.addEventListener('click', async () => {
        try {
          const prefs = {
            verboseLogging: document.getElementById('web-verbose-logging').checked === true
          };
          
          const res = await fetch(API_BASE + '/api/preferences', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(prefs)
          });
          
          const result = await res.json();
          if (result.success) {
            showStatus('Logging settings saved', false);
          } else {
            showStatus('Failed to save logging settings: ' + (result.error || 'Unknown error'), true);
          }
        } catch (error) {
          showStatus('Failed to save logging settings: ' + error.message, true);
        }
      });
    }
    
    // Backup status polling
    let webBackupStatusInterval = null;
    
    function startWebBackupStatusPolling() {
      stopWebBackupStatusPolling();
      
      updateWebBackupStatus();
      webBackupStatusInterval = setInterval(updateWebBackupStatus, 5000);
    }
    
    function stopWebBackupStatusPolling() {
      if (webBackupStatusInterval) {
        clearInterval(webBackupStatusInterval);
        webBackupStatusInterval = null;
      }
    }
    
    async function updateWebBackupStatus() {
      try {
        const response = await fetch(API_BASE + '/api/backup-status');
        if (!response.ok) {
          throw new Error('Failed to fetch backup status');
        }
        const data = await response.json();

        // Normalize into { ip -> status } and refresh the badges
        webBackupStatusByIp = {};
        if (data && Array.isArray(data.backups)) {
          data.backups.forEach((b) => {
            const ip = String(b?.ip || '').trim();
            if (!ip) return;
            webBackupStatusByIp[ip] = b?.status || null;
          });
        }
        refreshWebBackupStatusBadges();
      } catch (error) {
        console.error('Failed to update backup status:', error);
      }
    }
    
    // Update visibility in real-time when settings change
    document.getElementById('stagetimer-visible').addEventListener('change', () => {
      stagetimerVisible = document.getElementById('stagetimer-visible').checked;
      updateStagetimerVisibility();
    });
    
    document.getElementById('stagetimer-enabled').addEventListener('change', () => {
      stagetimerEnabled = document.getElementById('stagetimer-enabled').checked;
      updateStagetimerVisibility();
    });
    
    document.getElementById('stagetimer-api-key').addEventListener('input', () => {
      updateStagetimerVisibility();
    });
    
    document.getElementById('stagetimer-room-id').addEventListener('input', () => {
      updateStagetimerVisibility();
    });
    
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const data = { presetUrls: webGetPresetUrls() };
      fetch(API_BASE + '/api/presets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      })
        .then(res => {
          if (!res.ok) {
            throw new Error('HTTP error! status: ' + res.status);
          }
          return res.json();
        })
        .then(result => {
          if (result.success) {
            showStatus('Presets saved successfully!', false);
            // Reload presets to update the preset buttons
            fetch(API_BASE + '/api/presets')
              .then(res => res.json())
              .then(data => {
                createPresetButtons(data);
              })
              .catch(err => console.error('Failed to reload presets:', err));
          } else {
            showStatus('Failed to save: ' + (result.error || 'Unknown error'), true);
          }
        })
        .catch(err => {
          console.error('Fetch error:', err);
          let errorMsg = 'Failed to save presets: ' + err.message;
          if (err.message.includes('Failed to fetch')) {
            errorMsg += ' (Make sure the app is running and check network connection)';
          }
          showStatus(errorMsg, true);
        });
    });
    }
  </script>
</body>
</html>`;
      
      res.writeHead(200, {
        'Content-Type': 'text/html; charset=utf-8',
        'Cache-Control': 'no-store, no-cache, must-revalidate',
        'Pragma': 'no-cache'
      });
      res.end(html);
      return;
    }
    
    // Proxy API requests to the API server (so Web UI can work over port 80 only)
    if (req.url.startsWith('/api/')) {
      const prefs = loadPreferences();
      const apiPort = prefs.apiPort || DEFAULT_API_PORT;
      const apiReqPath = req.url.split('?')[0];
      const apiMethod = String(req.method || 'GET').toUpperCase();

      if (isWebUiRestrictedTunnelClient(req, prefs)) {
        const proxyForbidden =
          apiReqPath === '/api/preferences' ||
          apiReqPath === '/api/displays' ||
          apiReqPath === '/api/debug/preferences' ||
          (apiReqPath === '/api/stagetimer-settings' && apiMethod === 'POST') ||
          (apiReqPath === '/api/presets' && apiMethod === 'POST');
        if (proxyForbidden) {
          res.writeHead(403, {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type'
          });
          res.end(JSON.stringify({
            success: false,
            error: 'This action is not available on the shared link. Open the Web UI on the local network for Settings.'
          }));
          return;
        }
      }
      
      // Forward the request to the API server
      const apiReq = http.request({
        hostname: '127.0.0.1',
        port: apiPort,
        path: req.url,
        method: req.method,
        headers: {
          'Content-Type': req.headers['content-type'] || 'application/json'
        }
      }, (apiRes) => {
        // Copy response headers
        res.writeHead(apiRes.statusCode, {
          'Content-Type': apiRes.headers['content-type'] || 'application/json',
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type'
        });
        
        // Pipe the response
        apiRes.pipe(res);
      });
      
      apiReq.on('error', (err) => {
        console.error('[Web UI] Proxy error:', err);
        res.writeHead(502, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ 
          success: false, 
          error: 'Cannot connect to API server: ' + err.message 
        }));
      });
      
      // Forward request body if present
      req.pipe(apiReq);
      return;
    }
    
    // 404 for other routes
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not found');
  };

  const creds = getWebUiHttpsCredentials();
  const prefs = loadPreferences();
  let webUiPort = prefs.webUiPort || DEFAULT_WEB_UI_PORT;
  if (creds && webUiPort === 80) {
    webUiPort = DEFAULT_WEB_UI_HTTPS_PORT;
  }
  webUiServerUsesHttps = !!creds;
  if (creds) {
    webUiServer = https.createServer(creds, requestHandler);
  } else {
    webUiServer = http.createServer(requestHandler);
  }

  const protocol = creds ? 'https' : 'http';
  currentWebUiPort = webUiPort;
  webUiServer.listen(webUiPort, '0.0.0.0', () => {
    console.log(`[Web UI] Server listening on ${protocol}://0.0.0.0:${webUiPort}`);
    startCloudflaredTunnel();
  }).on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`[Web UI] Port ${webUiPort} is already in use`);
      if (process.env.GSO_README_CAPTURE !== '1') {
        dialog.showErrorBox(
          'Port Already in Use',
          `Port ${webUiPort} is already in use. Another instance of Google Slides Opener may be running.\n\nPlease quit the other instance or change the Web UI port in settings.`
        );
      }
      // Don't exit the app, but the server won't start
    } else {
      console.error('[Web UI] Server error:', err);
      if (process.env.GSO_README_CAPTURE !== '1') {
        dialog.showErrorBox(
          'Server Error',
          `Failed to start Web UI server: ${err.message}`
        );
      }
    }
  });
}

app.whenReady().then(() => {
  if (process.env.GSO_README_CAPTURE === '1') {
    try {
      readmeCapturePrefsBackup = JSON.parse(JSON.stringify(loadPreferences()));
      savePreferences({ ...readmeCapturePrefsBackup, webUiPort: 8765 });
    } catch (e) {
      logError('[readme-capture] could not apply capture Web UI port:', e);
    }
  }
  setupGoogleSessionEncoding();
  createWindow();
  startHttpServer();
  startWebUiServer();
  // Quick Tunnel starts from Web UI listen callback when cloudflaredEnabled is set

  // Start backup status polling if in primary mode
  startBackupStatusPolling();

  scheduleReadmeScreenshotCapture();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('before-quit', () => {
  stopCloudflaredTunnel();
  if (httpServer) {
    console.log('[API] Shutting down HTTP server');
    httpServer.close();
  }
  if (webUiServer) {
    console.log('[Web UI] Shutting down web UI server');
    webUiServer.close();
  }
});
