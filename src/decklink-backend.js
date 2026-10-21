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
 * `pixelFormat` is 'bgra' or 'yuv'. A mode the device rejects throws an error
 * matching /display mode is not supported/i, which openOutputNegotiated uses
 * to try the next format.
 */

const { PIXEL_FORMAT_PREFERENCE } = require('./decklink-modes');

// Only a rejected mode is worth retrying in another pixel format. Anything else
// (device busy, no access) would just fail again and a retry would hide the cause.
const MODE_REJECTED = /display mode is not supported/i;

const BACKENDS = {
  macadam: () => require('./backends/macadam-backend').createMacadamBackend(),
};

function createBackend(name = process.env.DECKLINK_BACKEND || 'macadam') {
  const factory = BACKENDS[name];
  if (!factory) throw new Error(`unknown DeckLink backend: ${name}`);
  return factory();
}

// Opens an output, trying each pixel format in turn until the device accepts one.
// Resolves to { output, pixelFormat }.
async function openOutputNegotiated(backend, opts, formats = PIXEL_FORMAT_PREFERENCE) {
  let lastError = null;
  for (const pixelFormat of formats) {
    try {
      const output = await backend.openOutput({ ...opts, pixelFormat });
      return { output, pixelFormat };
    } catch (e) {
      if (!MODE_REJECTED.test(e.message)) throw e;
      lastError = e;
    }
  }
  throw new Error(`${lastError.message} (tried pixel formats: ${formats.join(', ')})`);
}

module.exports = { createBackend, openOutputNegotiated };
