'use strict';

const { execFileSync } = require('child_process');

const DISPLAY_MODES = {
  '1080p2997': { bmdMode: 'bmdModeHD1080p2997', width: 1920, height: 1080, fps: 29.97 },
  '1080p25':   { bmdMode: 'bmdModeHD1080p25',   width: 1920, height: 1080, fps: 25   },
  '1080p30':   { bmdMode: 'bmdModeHD1080p30',   width: 1920, height: 1080, fps: 30   },
  '1080i5994': { bmdMode: 'bmdModeHD1080i5994', width: 1920, height: 1080, fps: 29.97 },
  '720p5994':  { bmdMode: 'bmdModeHD720p5994',  width: 1280, height: 720,  fps: 59.94 },
  '720p50':    { bmdMode: 'bmdModeHD720p50',     width: 1280, height: 720,  fps: 50   },
};

let _cachedProviderType = null;

async function detectProviderType() {
  if (_cachedProviderType !== null) return _cachedProviderType;

  try {
    const macadam = require('macadam');
    void macadam.bmdModeHD1080p2997;
    _cachedProviderType = 'macadam';
    return _cachedProviderType;
  } catch (e) {}

  try {
    execFileSync('ffmpeg', ['-hide_banner', '-f', 'decklink', '-list_devices', '1', '-i', 'dummy'],
      { stdio: 'pipe', timeout: 3000 });
    _cachedProviderType = 'ffmpeg';
    return _cachedProviderType;
  } catch (e) {
    if (e.code !== 'ENOENT') {
      // ffmpeg exists but errored (expected with dummy input); decklink support is present
      _cachedProviderType = 'ffmpeg';
      return _cachedProviderType;
    }
  }

  console.warn('[decklink] No DeckLink output provider found (macadam and ffmpeg unavailable)');
  _cachedProviderType = 'unavailable';
  return _cachedProviderType;
}

function createProvider(type) {
  return type === 'macadam' ? new MacadamProvider() : new FfmpegProvider();
}

class MacadamProvider {
  constructor() {
    this._macadam = null;
    this._playback = null;
    this._busy = false;
    this._mode = null;
  }

  async start(deviceIndex, displayModeKey) {
    this._macadam = require('macadam');
    const mode = DISPLAY_MODES[displayModeKey];
    this._playback = await this._macadam.playback({
      deviceIndex,
      displayMode: this._macadam[mode.bmdMode],
      pixelFormat: this._macadam.bmdFormat8BitBGRA
    });
    this._mode = mode;
  }

  pushFrame(bgraBuffer) {
    if (this._busy || !this._playback) return;
    this._busy = true;
    this._playback.frame(bgraBuffer, () => { this._busy = false; });
  }

  stop() {
    try { if (this._playback) this._playback.stop(); } catch (e) {}
    this._playback = null;
  }
}

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
      deviceArg
    ], { stdio: ['pipe', 'ignore', 'ignore'] });
    this._proc.on('error', () => {});
  }

  pushFrame(bgraBuffer) {
    if (this._busy || !this._proc || !this._proc.stdin.writable) return;
    this._busy = true;
    this._proc.stdin.write(bgraBuffer, () => { this._busy = false; });
  }

  stop() {
    try { if (this._proc) this._proc.kill('SIGTERM'); } catch (e) {}
    this._proc = null;
  }
}

class OutputController {
  constructor() {
    this._getWindow = null;
    this._provider = null;
    this._mode = null;
    this._blackFrame = null;
    this._lastGoodFrame = null;
    this._lastGoodTime = 0;
    this._timer = null;
  }

  start(getWindow, provider, displayModeKey) {
    const mode = DISPLAY_MODES[displayModeKey];
    this._getWindow = getWindow;
    this._provider = provider;
    this._mode = mode;
    this._blackFrame = Buffer.alloc(mode.width * mode.height * 4, 0);
    this._lastGoodFrame = null;
    this._lastGoodTime = 0;
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
      const image = await win.webContents.capturePage();
      const resized = image.resize({ width: this._mode.width, height: this._mode.height });
      const buf = resized.toBitmap();
      this._provider.pushFrame(buf);
      this._lastGoodFrame = buf;
      this._lastGoodTime = Date.now();
    } catch (e) {
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

class DecklinkOutputManagerClass {
  constructor() {
    this._providerType = 'unavailable';
    this._slidesProvider = null;
    this._notesProvider = null;
    this._slidesController = null;
    this._notesController = null;
    this._getSlidesWindow = null;
    this._getNotesWindow = null;
  }

  async init(getSlidesWindow, getNotesWindow, prefs) {
    this._getSlidesWindow = getSlidesWindow;
    this._getNotesWindow = getNotesWindow;
    await this._applyPrefs(prefs);
  }

  async reconfigure(prefs) {
    await this._stopAll();
    await this._applyPrefs(prefs);
  }

  async _applyPrefs(prefs) {
    const deckPrefs = prefs && prefs.decklink;
    if (!deckPrefs) return;

    this._providerType = await detectProviderType();
    if (this._providerType === 'unavailable') return;

    if (deckPrefs.slides && deckPrefs.slides.enabled) {
      try {
        const provider = createProvider(this._providerType);
        await provider.start(deckPrefs.slides.deviceIndex || 0, deckPrefs.slides.displayMode || '1080p2997');
        const controller = new OutputController();
        controller.start(this._getSlidesWindow, provider, deckPrefs.slides.displayMode || '1080p2997');
        this._slidesProvider = provider;
        this._slidesController = controller;
      } catch (e) {
        console.error('[decklink] Failed to start slides output:', e.message);
      }
    }

    if (deckPrefs.notes && deckPrefs.notes.enabled) {
      try {
        const provider = createProvider(this._providerType);
        await provider.start(deckPrefs.notes.deviceIndex || 1, deckPrefs.notes.displayMode || '1080p2997');
        const controller = new OutputController();
        controller.start(this._getNotesWindow, provider, deckPrefs.notes.displayMode || '1080p2997');
        this._notesProvider = provider;
        this._notesController = controller;
      } catch (e) {
        console.error('[decklink] Failed to start notes output:', e.message);
      }
    }
  }

  async _stopAll() {
    if (this._slidesController) { this._slidesController.stop(); this._slidesController = null; }
    if (this._notesController) { this._notesController.stop(); this._notesController = null; }
    if (this._slidesProvider) { this._slidesProvider.stop(); this._slidesProvider = null; }
    if (this._notesProvider) { this._notesProvider.stop(); this._notesProvider = null; }
  }

  async shutdown() { await this._stopAll(); }

  getStatus() {
    return {
      providerType: this._providerType,
      slides: { active: this._slidesController !== null },
      notes:  { active: this._notesController !== null }
    };
  }

  async getDevices() {
    try {
      const macadam = require('macadam');
      const info = await macadam.deviceInfo();
      return info.map((d, i) => ({ index: i, name: d.displayName || `DeckLink ${i}` }));
    } catch (e) {
      return [];
    }
  }
}

module.exports = { DecklinkOutputManager: new DecklinkOutputManagerClass() };
