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

const STATS_INTERVAL_MS = 5000;
// Electron's default resize quality is the slowest; 'good' is plenty for
// downscaling a capture to the output size. Env override is for A/B timing.
const RESIZE_QUALITY = process.env.DECKLINK_RESIZE_QUALITY || 'good';
const WORKER_SCRIPT = path.join(__dirname, 'macadam-worker.js');
const FFMPEG_PROBE_TIMEOUT_MS = 5000;
// ffmpeg exits within a moment if the device or mode is rejected; a process
// still alive after this long is treated as started.
const FFMPEG_STARTUP_GRACE_MS = 1500;
const FFMPEG_STDERR_TAIL = 600;

// pixelFormat 'yuv' marks modes the hardware only accepts as 8-bit YUV (e.g.
// 1080p50/59.94/60 on the UltraStudio Monitor 3G); the worker converts BGRA.
const DISPLAY_MODES = {
  '1080p5994': { bmdMode: 'bmdModeHD1080p5994', width: 1920, height: 1080, fps: 59.94, pixelFormat: 'yuv' },
  '1080p60':   { bmdMode: 'bmdModeHD1080p6000', width: 1920, height: 1080, fps: 60, pixelFormat: 'yuv' },
  '1080p50':   { bmdMode: 'bmdModeHD1080p50',   width: 1920, height: 1080, fps: 50, pixelFormat: 'yuv' },
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


// ── FFmpeg helpers (pure, unit tested) ────────────────────────────────────────

// `ffmpeg -devices` lists muxing devices as e.g. " E  decklink  Blackmagic DeckLink output".
function ffmpegListsDecklink(devicesOutput) {
  return /^\s*[D ]E\s+decklink\b/m.test(String(devicesOutput || ''));
}

// `ffmpeg -sinks decklink` prints a header, then one "  <name> [<description>]"
// line per output device (a leading "*" replaces the indent for the default).
function parseFfmpegSinks(sinksOutput) {
  const names = [];
  for (const line of String(sinksOutput || '').split(/\r?\n/)) {
    if (!/^[\s*]+\S/.test(line)) continue;         // device lines start with a space, or * for the default
    const name = line.trim().replace(/^\*\s*/, '').replace(/\s*\[[^\]]*\]\s*$/, '');
    if (name) names.push(name);
  }
  return names;
}

// The DeckLink muxer matches modes by exact rational: 59.94 must be 60000/1001.
function ffmpegFrameRate(fps) {
  return Number.isInteger(fps) ? String(fps) : `${Math.round(fps) * 1000}/1001`;
}

function buildFfmpegArgs(mode, deviceName) {
  return [
    '-hide_banner',
    '-f', 'rawvideo',
    '-pixel_format', 'bgra',
    '-video_size', `${mode.width}x${mode.height}`,
    '-framerate', ffmpegFrameRate(mode.fps),
    '-i', 'pipe:0',
    '-f', 'decklink',
    '-pix_fmt', 'uyvy422',
    deviceName,
  ];
}

// Returns { available, error }. `run` is injectable for tests.
function detectFfmpeg(run = execFileSync) {
  try {
    const out = run('ffmpeg', ['-hide_banner', '-devices'],
      { stdio: 'pipe', timeout: FFMPEG_PROBE_TIMEOUT_MS, encoding: 'utf8' });
    if (ffmpegListsDecklink(out)) return { available: true, error: null };
    return { available: false, error: 'ffmpeg was built without DeckLink support' };
  } catch (e) {
    return { available: false, error: e.code === 'ENOENT' ? 'ffmpeg not found' : e.message };
  }
}

function ffmpegDeviceName(deviceIndex, run = execFileSync) {
  try {
    const out = run('ffmpeg', ['-hide_banner', '-sinks', 'decklink'],
      { stdio: 'pipe', timeout: FFMPEG_PROBE_TIMEOUT_MS, encoding: 'utf8' });
    const names = parseFfmpegSinks(out);
    if (names[deviceIndex]) return names[deviceIndex];
  } catch (_) { /* fall through to the generic name */ }
  return deviceIndex === 0 ? 'DeckLink Output' : `DeckLink Output ${deviceIndex + 1}`;
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

  // FFmpeg fallback: only usable if this ffmpeg build actually has DeckLink output.
  const ff = detectFfmpeg();
  if (ff.available) {
    _cachedProviderType = 'ffmpeg';
    console.log('[decklink] Provider: ffmpeg');
    return _cachedProviderType;
  }
  _detectionErrors.ffmpeg = ff.error;
  console.warn('[decklink] ffmpeg unavailable:', ff.error);

  console.warn('[decklink] No provider found. macadam:', _detectionErrors.macadam,
    '| ffmpeg:', _detectionErrors.ffmpeg);
  _cachedProviderType = 'unavailable';
  return _cachedProviderType;
}

function createProvider(type) {
  return type === 'macadam' ? new MacadamProvider() : new FfmpegProvider();
}

// ── MacadamProvider — persistent worker per output ────────────────────────────
// macadam.playback() crashes with SIGSEGV when called in the Electron main
// process. Running it in a fork isolates the crash so the app stays alive.

class MacadamProvider {
  constructor() {
    this._worker  = null;
    this._msgId   = 0;
    this._pending = new Map();
    this._busy    = false;
    this._mode    = null;
    this._pbId    = null;
    this._timings = { frames: 0, roundTripMs: 0, convertMs: 0, displayMs: 0 };
  }

  // Per-frame worker timings since the last call (for the periodic stats line).
  takeTimings() {
    const t = this._timings;
    this._timings = { frames: 0, roundTripMs: 0, convertMs: 0, displayMs: 0 };
    return t.frames > 0 ? t : null;
  }

  _spawnWorker() {
    const w = fork(WORKER_SCRIPT, [], { serialization: 'advanced', silent: true });

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
      console.error(`[decklink] macadam worker exited (code=${code} signal=${signal})`);
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
      await this._send('start', {
        deviceIndex,
        bmdMode: mode.bmdMode,
        pixelFormat: mode.pixelFormat || 'bgra',
        width: mode.width,
        height: mode.height,
        playbackId: this._pbId,
      });
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

// ── FfmpegProvider ────────────────────────────────────────────────────────────

class FfmpegProvider {
  // `deps` lets tests inject a fake spawn / device lookup / startup grace.
  constructor(deps = {}) {
    this._spawn   = deps.spawn || require('child_process').spawn;
    this._nameFor = deps.deviceName || ffmpegDeviceName;
    this._graceMs = deps.graceMs !== undefined ? deps.graceMs : FFMPEG_STARTUP_GRACE_MS;
    this._proc    = null;
    this._busy    = false;
    this._mode    = null;
    this._stderr  = '';
  }

  async start(deviceIndex, displayModeKey) {
    const mode = DISPLAY_MODES[displayModeKey];
    this._mode = mode;
    this._stderr = '';
    const proc = this._spawn('ffmpeg', buildFfmpegArgs(mode, this._nameFor(deviceIndex)),
      { stdio: ['pipe', 'ignore', 'pipe'] });
    this._proc = proc;

    if (proc.stderr) {
      proc.stderr.on('data', (chunk) => {
        this._stderr = (this._stderr + chunk.toString()).slice(-FFMPEG_STDERR_TAIL);
      });
    }
    // ffmpeg dying closes the pipe; an unhandled EPIPE here would take down the app.
    if (proc.stdin) proc.stdin.on('error', () => {});

    await new Promise((resolve, reject) => {
      const timer = setTimeout(resolve, this._graceMs);
      const fail = (msg) => {
        clearTimeout(timer);
        this._proc = null;
        reject(new Error(msg));
      };
      proc.once('error', (e) => fail(`ffmpeg failed to start: ${e.message}`));
      proc.once('exit', (code, signal) => fail(
        `ffmpeg exited (code=${code} signal=${signal})` +
        (this._stderr ? `: ${this._stderr.trim().split(/\r?\n/).pop()}` : '')));
    });

    // Later exit: drop the process so pushFrame stops writing into a dead pipe.
    proc.on('exit', (code, signal) => {
      console.error(`[decklink] ffmpeg exited (code=${code} signal=${signal})`);
      if (this._proc === proc) this._proc = null;
      this._busy = false;
    });
  }

  pushFrame(bgraBuffer) {
    if (this._busy || !this._proc || !this._proc.stdin || !this._proc.stdin.writable) return false;
    this._busy = true;
    this._proc.stdin.write(bgraBuffer, () => { this._busy = false; });
    return true;
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
        const provider = createProvider(this._providerType);
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
        const provider = createProvider(this._providerType);
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

module.exports = {
  DecklinkOutputManager: new DecklinkOutputManagerClass(),
  DISPLAY_MODES,
  OutputController,
  FfmpegProvider,
  ffmpegListsDecklink,
  parseFfmpegSinks,
  ffmpegFrameRate,
  buildFfmpegArgs,
  detectFfmpeg,
};
