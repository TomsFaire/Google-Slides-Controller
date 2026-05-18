'use strict';

/**
 * decklink-output.js
 *
 * All macadam (native DeckLink SDK) calls are made inside a child_process
 * worker (macadam-worker.js).  If the DeckLink SDK crashes in a native
 * thread the worker process dies, but the Electron main process survives.
 *
 * FFmpeg is still supported as a fallback for machines without macadam.
 */

const path = require('path');
const { fork, execFileSync } = require('child_process');

const WORKER_SCRIPT = path.join(__dirname, 'macadam-worker.js');

const DISPLAY_MODES = {
  '1080p5994': { bmdMode: 'bmdModeHD1080p5994', width: 1920, height: 1080, fps: 59.94 },
  '1080p60':   { bmdMode: 'bmdModeHD1080p60',   width: 1920, height: 1080, fps: 60   },
  '1080p50':   { bmdMode: 'bmdModeHD1080p50',   width: 1920, height: 1080, fps: 50   },
  '1080p2997': { bmdMode: 'bmdModeHD1080p2997', width: 1920, height: 1080, fps: 29.97 },
  '1080p25':   { bmdMode: 'bmdModeHD1080p25',   width: 1920, height: 1080, fps: 25   },
  '1080p30':   { bmdMode: 'bmdModeHD1080p30',   width: 1920, height: 1080, fps: 30   },
  '1080i5994': { bmdMode: 'bmdModeHD1080i5994', width: 1920, height: 1080, fps: 29.97 },
  '720p5994':  { bmdMode: 'bmdModeHD720p5994',  width: 1280, height: 720,  fps: 59.94 },
  '720p50':    { bmdMode: 'bmdModeHD720p50',     width: 1280, height: 720,  fps: 50   },
};

let _cachedProviderType = null;
let _detectionErrors = {};

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
    await workerRpc('probe');
    _cachedProviderType = 'macadam';
    console.log('[decklink] Provider: macadam (worker)');
    return _cachedProviderType;
  } catch (e) {
    _detectionErrors.macadam = e.message;
    console.warn('[decklink] macadam probe failed:', e.message);
  }

  // FFmpeg fallback
  try {
    execFileSync(
      'ffmpeg',
      ['-hide_banner', '-f', 'decklink', '-list_devices', '1', '-i', 'dummy'],
      { stdio: 'pipe', timeout: 3000 }
    );
    _cachedProviderType = 'ffmpeg';
    console.log('[decklink] Provider: ffmpeg');
    return _cachedProviderType;
  } catch (e) {
    if (e.code !== 'ENOENT') {
      _cachedProviderType = 'ffmpeg';
      console.log('[decklink] Provider: ffmpeg');
      return _cachedProviderType;
    }
    _detectionErrors.ffmpeg = e.message;
    console.warn('[decklink] ffmpeg unavailable:', e.message);
  }

  console.warn('[decklink] No provider found. macadam:', _detectionErrors.macadam,
    '| ffmpeg:', _detectionErrors.ffmpeg);
  _cachedProviderType = 'unavailable';
  return _cachedProviderType;
}

function createProvider(type) {
  return type === 'macadam' ? new MacadamProvider() : new FfmpegProvider();
}

// ── MacadamProvider — runs macadam directly in the main process ───────────────

class MacadamProvider {
  constructor() {
    this._macadam  = null;
    this._playback = null;
    this._busy     = false;
    this._mode     = null;
  }

  async start(deviceIndex, displayModeKey) {
    this._macadam = require('macadam');
    const mode = DISPLAY_MODES[displayModeKey];
    this._mode = mode;
    this._playback = await this._macadam.playback({
      deviceIndex,
      displayMode: this._macadam[mode.bmdMode],
      pixelFormat: this._macadam.bmdFormat8BitBGRA,
    });
  }

  pushFrame(bgraBuffer) {
    if (this._busy || !this._playback) return;
    this._busy = true;
    this._playback.frame(bgraBuffer, () => { this._busy = false; });
  }

  stop() {
    try { if (this._playback) this._playback.stop(); } catch (_) {}
    this._playback = null;
  }
}

// ── FfmpegProvider ────────────────────────────────────────────────────────────

class FfmpegProvider {
  constructor() {
    this._proc = null;
    this._busy = false;
    this._mode = null;
  }

  async start(deviceIndex, displayModeKey) {
    const mode = DISPLAY_MODES[displayModeKey];
    this._mode = mode;
    const deviceArg = deviceIndex === 0 ? 'DeckLink Output' : `DeckLink Output ${deviceIndex + 1}`;
    this._proc = require('child_process').spawn('ffmpeg', [
      '-hide_banner',
      '-f', 'rawvideo',
      '-pixel_format', 'bgra',
      '-video_size', `${mode.width}x${mode.height}`,
      '-framerate', String(mode.fps),
      '-i', 'pipe:0',
      '-f', 'decklink',
      '-pix_fmt', 'uyvy422',
      deviceArg,
    ], { stdio: ['pipe', 'ignore', 'ignore'] });
    this._proc.on('error', () => {});
  }

  pushFrame(bgraBuffer) {
    if (this._busy || !this._proc || !this._proc.stdin.writable) return;
    this._busy = true;
    this._proc.stdin.write(bgraBuffer, () => { this._busy = false; });
  }

  stop() {
    try { if (this._proc) this._proc.kill('SIGTERM'); } catch (_) {}
    this._proc = null;
  }
}

// ── OutputController ──────────────────────────────────────────────────────────

class OutputController {
  constructor() {
    this._getWindow     = null;
    this._provider      = null;
    this._mode          = null;
    this._blackFrame    = null;
    this._lastGoodFrame = null;
    this._lastGoodTime  = 0;
    this._timer         = null;
  }

  start(getWindow, provider, displayModeKey) {
    const mode = DISPLAY_MODES[displayModeKey];
    this._getWindow  = getWindow;
    this._provider   = provider;
    this._mode       = mode;
    this._blackFrame = Buffer.alloc(mode.width * mode.height * 4, 0);
    const intervalMs = Math.round(1000 / mode.fps);
    this._timer = setInterval(() => this._tick(), intervalMs);
  }

  async _tick() {
    const win = this._getWindow();
    if (!win || win.isDestroyed()) {
      this._provider.pushFrame(this._blackFrame);
      return;
    }
    try {
      const image   = await win.webContents.capturePage();
      const resized = image.resize({ width: this._mode.width, height: this._mode.height });
      const buf     = resized.toBitmap();
      this._provider.pushFrame(buf);
      this._lastGoodFrame = buf;
      this._lastGoodTime  = Date.now();
    } catch (_) {
      const buf = (this._lastGoodFrame && (Date.now() - this._lastGoodTime) < 2000)
        ? this._lastGoodFrame
        : this._blackFrame;
      this._provider.pushFrame(buf);
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
        const provider = createProvider(this._providerType);
        await provider.start(
          deckPrefs.slides.deviceIndex || 0,
          deckPrefs.slides.displayMode || '1080p5994'
        );
        const controller = new OutputController();
        controller.start(
          this._getSlidesWindow, provider,
          deckPrefs.slides.displayMode || '1080p5994'
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
        const provider = createProvider(this._providerType);
        await provider.start(
          deckPrefs.notes.deviceIndex || 1,
          deckPrefs.notes.displayMode || '1080p5994'
        );
        const controller = new OutputController();
        controller.start(
          this._getNotesWindow, provider,
          deckPrefs.notes.displayMode || '1080p5994'
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
      detectionErrors: _detectionErrors,
      slides: { active: this._slidesController !== null, error: this._slidesError || null },
      notes:  { active: this._notesController  !== null, error: this._notesError  || null },
    };
  }

  async getDevices() {
    try {
      const msg = await workerRpc('get_devices');
      return msg.devices || [];
    } catch (e) {
      console.warn('[decklink] getDevices failed:', e.message);
      return [];
    }
  }
}

module.exports = { DecklinkOutputManager: new DecklinkOutputManagerClass() };
