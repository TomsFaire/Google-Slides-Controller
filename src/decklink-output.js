'use strict';

/**
 * decklink-output.js
 *
 * All macadam (native DeckLink SDK) calls are made inside a child_process
 * worker (decklink-worker.js).  If the DeckLink SDK crashes in a native
 * thread the worker process dies, but the Electron main process survives.
 */

const path = require('path');
const { fork } = require('child_process');

const STATS_INTERVAL_MS = 5000;
// Electron's default resize quality is the slowest; 'good' is plenty for
// downscaling a capture to the output size. Env override is for A/B timing.
const RESIZE_QUALITY = process.env.DECKLINK_RESIZE_QUALITY || 'good';
const { DISPLAY_MODES, filterOutputDevices } = require('./decklink-modes');

const WORKER_SCRIPT = path.join(__dirname, 'decklink-worker.js');

let _cachedProviderType = null;
let _detectionErrors = {};
let _driverVersion = null;

// ── Worker helpers ────────────────────────────────────────────────────────────

/**
 * Spawn a short-lived worker, send one command, return the response or throw.
 */
function workerRpc(cmd, payload = {}, timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const settle = (fn, val) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      try { worker.kill(); } catch (_) {}
      fn(val);
    };

    const worker = fork(WORKER_SCRIPT, [], {
      serialization: 'advanced',
      silent: true,
    });

    const timer = setTimeout(
      () => settle(reject, new Error('macadam worker timeout')),
      timeoutMs
    );

    worker.once('message', (msg) => {
      if (msg.cmd === 'error') settle(reject, new Error(msg.error));
      else settle(resolve, msg);
    });
    worker.on('error', (e) => settle(reject, e));
    worker.on('exit', (code, signal) => {
      if (!settled)
        settle(reject, new Error(
          `macadam worker exited unexpectedly (code=${code} signal=${signal})`
        ));
    });

    worker.send({ id: 1, cmd, ...payload });
  });
}


// ── Provider detection ────────────────────────────────────────────────────────

async function detectProviderType() {
  if (_cachedProviderType !== null) return _cachedProviderType;

  // Try macadam in an isolated worker — a native crash won't kill us.
  try {
    const probe = await workerRpc('probe');
    _cachedProviderType = 'macadam';
    _driverVersion = probe.driverVersion || null;
    console.log('[decklink] Provider: macadam (worker)' +
      (_driverVersion ? `, ${_driverVersion}` : ''));
    return _cachedProviderType;
  } catch (e) {
    _detectionErrors.macadam = e.message;
    console.warn('[decklink] macadam probe failed:', e.message);
  }

  console.warn('[decklink] No provider found. macadam:', _detectionErrors.macadam);
  _cachedProviderType = 'unavailable';
  return _cachedProviderType;
}

function createProvider() {
  return new WorkerProvider();
}

// ── WorkerProvider — persistent worker per output ─────────────────────────────
// macadam.playback() crashes with SIGSEGV when called in the Electron main
// process. Running it in a fork isolates the crash so the app stays alive.

class WorkerProvider {
  // `deps.fork` lets tests inject a fake worker process.
  constructor(deps = {}) {
    this._fork    = deps.fork || fork;
    this._worker  = null;
    this._msgId   = 0;
    this._pending = new Map();
    this._busy    = false;
    this._mode    = null;
    this._pbId    = null;
    this._timings = { frames: 0, roundTripMs: 0, convertMs: 0, displayMs: 0 };
    this.pixelFormat = null; // what the device accepted, set by start()
  }

  // Per-frame worker timings since the last call (for the periodic stats line).
  takeTimings() {
    const t = this._timings;
    this._timings = { frames: 0, roundTripMs: 0, convertMs: 0, displayMs: 0 };
    return t.frames > 0 ? t : null;
  }

  _spawnWorker() {
    const w = this._fork(WORKER_SCRIPT, [], { serialization: 'advanced', silent: true });

    w.on('message', (msg) => {
      const p = this._pending.get(msg.id);
      if (!p) return;
      this._pending.delete(msg.id);
      if (msg.cmd === 'error') p.reject(new Error(msg.error));
      else p.resolve(msg);
    });

    // Log stderr from worker so we can see native crash messages.
    if (w.stderr) {
      w.stderr.on('data', (chunk) => {
        console.error('[decklink] worker stderr:', chunk.toString().trimEnd());
      });
    }

    w.on('error', (e) => {
      console.error('[decklink] worker error event:', e.message);
    });

    w.on('exit', (code, signal) => {
      console.error(`[decklink] worker exited (code=${code} signal=${signal})`);
      this._worker = null;
      this._busy   = false;
      for (const [, p] of this._pending) p.reject(new Error(`worker exited (code=${code} signal=${signal})`));
      this._pending.clear();
    });

    return w;
  }

  _send(cmd, extra = {}) {
    return new Promise((resolve, reject) => {
      const id = ++this._msgId;
      this._pending.set(id, { resolve, reject });
      try {
        this._worker.send({ id, cmd, ...extra });
      } catch (e) {
        this._pending.delete(id);
        reject(e);
      }
    });
  }

  async start(deviceIndex, displayModeKey) {
    const mode = DISPLAY_MODES[displayModeKey];
    this._mode = mode;
    this._pbId = `pb-${deviceIndex}-${Date.now()}`;
    this._worker = this._spawnWorker();
    try {
      const reply = await this._send('start', {
        deviceIndex,
        bmdMode: mode.bmdMode,
        pixelFormats: mode.pixelFormats, // normally unset: the worker negotiates
        width: mode.width,
        height: mode.height,
        playbackId: this._pbId,
      });
      this.pixelFormat = reply.pixelFormat;
      console.log(`[decklink] output started: ${displayModeKey}, pixel format ${reply.pixelFormat}`);
    } catch (e) {
      // A failed start must not leave the worker alive: it can keep the
      // device claimed and lock out every later start attempt.
      const w = this._worker;
      this._worker = null;
      try { w.kill(); } catch (_) {}
      throw e;
    }
  }

  pushFrame(bgraBuffer) {
    if (this._busy || !this._worker) return false;
    this._busy = true;
    const id = ++this._msgId;
    const sentAt = Date.now();
    this._pending.set(id, {
      resolve: (msg) => {
        this._busy = false;
        if (msg && msg.timing) {
          const t = this._timings;
          t.frames++;
          t.roundTripMs += Date.now() - sentAt;
          t.convertMs   += msg.timing.convertMs;
          t.displayMs   += msg.timing.displayMs;
        }
      },
      reject:  () => { this._busy = false; },
    });
    try {
      const ab = bgraBuffer.buffer.slice(
        bgraBuffer.byteOffset,
        bgraBuffer.byteOffset + bgraBuffer.byteLength
      );
      this._worker.send({ id, cmd: 'frame', playbackId: this._pbId, data: ab });
      return true;
    } catch (_) {
      this._pending.delete(id);
      this._busy = false;
      return false;
    }
  }

  stop() {
    if (this._worker) {
      try { this._send('stop', { playbackId: this._pbId }).catch(() => {}); } catch (_) {}
      const w = this._worker;
      this._worker = null;
      setTimeout(() => { try { w.kill(); } catch (_) {} }, 600);
    }
  }
}

// ── OutputController ──────────────────────────────────────────────────────────

class OutputController {
  constructor() {
    this._getWindow     = null;
    this._provider      = null;
    this._mode          = null;
    this._label         = '';
    this._blackFrame    = null;
    this._lastGoodFrame = null;
    this._lastGoodTime  = 0;
    this._timer         = null;
    this._capturing     = false;
    this._resetStats();
  }

  _resetStats() {
    this._stats = {
      since: Date.now(), frames: 0, skippedTicks: 0, droppedPushes: 0,
      captureMs: 0, convertMs: 0,
    };
  }

  start(getWindow, provider, displayModeKey, label = '') {
    const mode = DISPLAY_MODES[displayModeKey];
    this._getWindow  = getWindow;
    this._provider   = provider;
    this._mode       = mode;
    this._label      = label;
    this._blackFrame = Buffer.alloc(mode.width * mode.height * 4, 0);
    const intervalMs = Math.round(1000 / mode.fps);
    this._timer = setInterval(() => this._tick(), intervalMs);
  }

  _push(buf) {
    if (this._provider.pushFrame(buf) === false) this._stats.droppedPushes++;
  }

  _logStats() {
    const s = this._stats;
    const elapsed = Date.now() - s.since;
    if (elapsed < STATS_INTERVAL_MS) return;
    if (s.frames > 0) {
      const w = this._provider.takeTimings && this._provider.takeTimings();
      const worker = w
        ? `, worker ${(w.frames * 1000 / elapsed).toFixed(1)} fps delivered: ` +
          `round-trip ${(w.roundTripMs / w.frames).toFixed(1)}ms ` +
          `(yuv ${(w.convertMs / w.frames).toFixed(1)}ms, displayFrame ${(w.displayMs / w.frames).toFixed(1)}ms)`
        : '';
      console.log(
        `[decklink] ${this._label}: ${(s.frames * 1000 / elapsed).toFixed(1)} fps captured, ` +
        `capture ${(s.captureMs / s.frames).toFixed(1)}ms, resize ${(s.convertMs / s.frames).toFixed(1)}ms, ` +
        `${s.skippedTicks} ticks skipped (capture busy), ${s.droppedPushes} frames dropped (device busy)${worker}`
      );
    }
    this._resetStats();
  }

  async _tick() {
    const win = this._getWindow();
    if (!win || win.isDestroyed()) {
      this._push(this._blackFrame);
      return;
    }
    // One capture at a time. Without this, ticks outpace capturePage() and the
    // queue grows, so the output falls further and further behind the window.
    if (this._capturing) {
      this._stats.skippedTicks++;
      return;
    }
    this._capturing = true;
    try {
      const t0      = Date.now();
      const image   = await win.webContents.capturePage();
      const t1      = Date.now();
      const { width, height } = this._mode;
      const size    = image.getSize();
      const resized = (size.width === width && size.height === height)
        ? image
        : image.resize({ width, height, quality: RESIZE_QUALITY });
      const buf     = resized.toBitmap();
      const t2      = Date.now();
      this._push(buf);
      this._lastGoodFrame = buf;
      this._lastGoodTime  = t2;
      this._stats.frames++;
      this._stats.captureMs += t1 - t0;
      this._stats.convertMs += t2 - t1;
    } catch (_) {
      const buf = (this._lastGoodFrame && (Date.now() - this._lastGoodTime) < 2000)
        ? this._lastGoodFrame
        : this._blackFrame;
      this._push(buf);
    } finally {
      this._capturing = false;
      this._logStats();
    }
  }

  stop() {
    clearInterval(this._timer);
    this._timer = null;
  }
}

// ── DecklinkOutputManager ─────────────────────────────────────────────────────

class DecklinkOutputManagerClass {
  constructor() {
    this._providerType     = 'unavailable';
    this._slidesError      = null;
    this._notesError       = null;
    this._slidesProvider   = null;
    this._notesProvider    = null;
    this._slidesController = null;
    this._notesController  = null;
    this._getSlidesWindow  = null;
    this._getNotesWindow   = null;
  }

  async init(getSlidesWindow, getNotesWindow, prefs) {
    this._getSlidesWindow = getSlidesWindow;
    this._getNotesWindow  = getNotesWindow;
    await this._applyPrefs(prefs);
  }

  async reconfigure(prefs) {
    await this._stopAll();
    _cachedProviderType = null;
    await this._applyPrefs(prefs);
  }

  async _applyPrefs(prefs) {
    const deckPrefs = prefs && prefs.decklink;
    if (!deckPrefs) return;

    this._providerType = await detectProviderType();
    if (this._providerType === 'unavailable') return;

    this._slidesError = null;
    this._notesError  = null;

    if (deckPrefs.slides && deckPrefs.slides.enabled) {
      try {
        const provider = createProvider();
        await provider.start(
          deckPrefs.slides.deviceIndex || 0,
          deckPrefs.slides.displayMode || '1080p5994'
        );
        const controller = new OutputController();
        controller.start(
          this._getSlidesWindow, provider,
          deckPrefs.slides.displayMode || '1080p5994', 'slides'
        );
        this._slidesProvider   = provider;
        this._slidesController = controller;
      } catch (e) {
        this._slidesError = e.message;
        console.error('[decklink] Failed to start slides output:', e.message);
      }
    }

    if (deckPrefs.notes && deckPrefs.notes.enabled) {
      try {
        const provider = createProvider();
        await provider.start(
          deckPrefs.notes.deviceIndex || 1,
          deckPrefs.notes.displayMode || '1080p5994'
        );
        const controller = new OutputController();
        controller.start(
          this._getNotesWindow, provider,
          deckPrefs.notes.displayMode || '1080p5994', 'notes'
        );
        this._notesProvider   = provider;
        this._notesController = controller;
      } catch (e) {
        this._notesError = e.message;
        console.error('[decklink] Failed to start notes output:', e.message);
      }
    }
  }

  async _stopAll() {
    if (this._slidesController) { this._slidesController.stop(); this._slidesController = null; }
    if (this._notesController)  { this._notesController.stop();  this._notesController  = null; }
    if (this._slidesProvider)   { this._slidesProvider.stop();   this._slidesProvider   = null; }
    if (this._notesProvider)    { this._notesProvider.stop();    this._notesProvider    = null; }
  }

  async shutdown() { await this._stopAll(); }

  getStatus() {
    return {
      providerType:    this._providerType,
      driverVersion:   _driverVersion,
      detectionErrors: _detectionErrors,
      slides: { active: this._slidesController !== null, error: this._slidesError || null,
                pixelFormat: this._slidesProvider ? this._slidesProvider.pixelFormat : null },
      notes:  { active: this._notesController  !== null, error: this._notesError  || null,
                pixelFormat: this._notesProvider ? this._notesProvider.pixelFormat : null },
    };
  }

  async getModes(deviceIndex) {
    try {
      const msg = await workerRpc('get_modes', { deviceIndex: Number(deviceIndex) || 0 });
      return msg.modes || [];
    } catch (e) {
      console.warn('[decklink] getModes failed:', e.message);
      return [];
    }
  }

  async getDevices() {
    try {
      const msg = await workerRpc('get_devices');
      return filterOutputDevices(msg.devices || []);
    } catch (e) {
      console.warn('[decklink] getDevices failed:', e.message);
      return [];
    }
  }
}

module.exports = {
  DecklinkOutputManager: new DecklinkOutputManagerClass(),
  DISPLAY_MODES,
  OutputController,
  WorkerProvider,
};
