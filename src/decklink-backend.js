'use strict';

/**
 * decklink-backend.js
 *
 * The only place that knows which native layer drives the DeckLink hardware.
 * A backend is:
 *   probe():       Promise<{ driverVersion: string | null }>
 *   listDevices(): Promise<Array<{ index, name, canOutput }>>
 *   openOutput({ deviceIndex, mode, width, height, pixelFormat }):
 *                  Promise<{ displayFrame(bgra): Promise<{ convertMs, displayMs }>,
 *                            close(): Promise<void> }>
 * `mode` is the BMD display-mode name (e.g. 'bmdModeHD1080p2997');
 * `pixelFormat` is 'bgra' or 'yuv'.
 */

const BACKENDS = {
  macadam: () => require('./backends/macadam-backend').createMacadamBackend(),
};

function createBackend(name = process.env.DECKLINK_BACKEND || 'macadam') {
  const factory = BACKENDS[name];
  if (!factory) throw new Error(`unknown DeckLink backend: ${name}`);
  return factory();
}

module.exports = { createBackend };
