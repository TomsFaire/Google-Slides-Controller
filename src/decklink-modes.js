'use strict';

/**
 * decklink-modes.js
 * Display modes the app offers, and the helpers that match them to what a
 * device actually supports.
 */

// Which pixel format a device takes for a mode is negotiated at start-up
// (see openOutputNegotiated), not listed here: the UltraStudio Express Monitor
// 3G, for example, accepts 1080p50/59.94/60 only as 8-bit YUV.
const DISPLAY_MODES = {
  '1080p5994': { bmdMode: 'bmdModeHD1080p5994', width: 1920, height: 1080, fps: 59.94 },
  '1080p60':   { bmdMode: 'bmdModeHD1080p6000', width: 1920, height: 1080, fps: 60    },
  '1080p50':   { bmdMode: 'bmdModeHD1080p50',   width: 1920, height: 1080, fps: 50    },
  '1080p2997': { bmdMode: 'bmdModeHD1080p2997', width: 1920, height: 1080, fps: 29.97 },
  '1080p25':   { bmdMode: 'bmdModeHD1080p25',   width: 1920, height: 1080, fps: 25    },
  '1080p30':   { bmdMode: 'bmdModeHD1080p30',   width: 1920, height: 1080, fps: 30    },
  '1080i5994': { bmdMode: 'bmdModeHD1080i5994', width: 1920, height: 1080, fps: 29.97 },
  '720p5994':  { bmdMode: 'bmdModeHD720p5994',  width: 1280, height: 720,  fps: 59.94 },
  '720p50':    { bmdMode: 'bmdModeHD720p50',    width: 1280, height: 720,  fps: 50    },
};

// BGRA needs no conversion, so it wins whenever the device accepts it.
const PIXEL_FORMAT_PREFERENCE = ['bgra', 'yuv'];

function choosePixelFormat(supported, { mode = 'this mode', deviceModes = [] } = {}) {
  const chosen = PIXEL_FORMAT_PREFERENCE.find((f) => supported.includes(f));
  if (chosen) return chosen;
  const offered = deviceModes.length ? deviceModes.join(', ') : 'none reported';
  throw new Error(`${mode} is not supported by this device (it supports: ${offered})`);
}

// Capture-only devices (recorders) cannot show a picture; `canOutput` unknown means keep.
function filterOutputDevices(devices) {
  return devices.filter((d) => d.canOutput !== false);
}

module.exports = { DISPLAY_MODES, PIXEL_FORMAT_PREFERENCE, choosePixelFormat, filterOutputDevices };
