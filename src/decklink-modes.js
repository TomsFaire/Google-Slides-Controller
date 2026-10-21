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
  '1080p5994': { bmdMode: 'bmdModeHD1080p5994', width: 1920, height: 1080, fps: 59.94, label: '1080p 59.94fps' },
  '1080p60':   { bmdMode: 'bmdModeHD1080p6000', width: 1920, height: 1080, fps: 60,    label: '1080p 60fps' },
  '1080p50':   { bmdMode: 'bmdModeHD1080p50',   width: 1920, height: 1080, fps: 50,    label: '1080p 50fps' },
  '1080p2997': { bmdMode: 'bmdModeHD1080p2997', width: 1920, height: 1080, fps: 29.97, label: '1080p 29.97fps' },
  '1080p25':   { bmdMode: 'bmdModeHD1080p25',   width: 1920, height: 1080, fps: 25,    label: '1080p 25fps' },
  '1080p30':   { bmdMode: 'bmdModeHD1080p30',   width: 1920, height: 1080, fps: 30,    label: '1080p 30fps' },
  '1080i5994': { bmdMode: 'bmdModeHD1080i5994', width: 1920, height: 1080, fps: 29.97, label: '1080i 59.94fps', interlaced: true },
  '720p5994':  { bmdMode: 'bmdModeHD720p5994',  width: 1280, height: 720,  fps: 59.94, label: '720p 59.94fps' },
  '720p50':    { bmdMode: 'bmdModeHD720p50',    width: 1280, height: 720,  fps: 50,    label: '720p 50fps' },
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

// The modes in DISPLAY_MODES are the ones the pipeline can drive (up to 1080p60,
// 8-bit). `matchDeviceModes` narrows them to what a device says it supports.
// `deviceModes` has macadam's shape: { name, width, height, frameRate: [duration, scale],
// videoModes: ['8-bit YUV', '8-bit BGRA', ...] }. Interlaced and progressive modes can
// share a frame rate (1080i59.94 and 1080p29.97 both report 29.97), so the name decides.
const DEVICE_PIXEL_FORMATS = { '8-bit BGRA': 'bgra', '8-bit YUV': 'yuv' };
const FPS_TOLERANCE = 0.01;

function isInterlacedName(name) {
  const m = /\d{3,4}\s*([ip])\b/i.exec(String(name || ''));
  return !!m && m[1].toLowerCase() === 'i';
}

function deviceFps(frameRate) {
  if (!Array.isArray(frameRate) || !frameRate[0]) return NaN;
  return frameRate[1] / frameRate[0];
}

function matchDeviceModes(deviceModes) {
  const list = Array.isArray(deviceModes) ? deviceModes : [];
  const result = [];
  for (const [key, known] of Object.entries(DISPLAY_MODES)) {
    const found = list.find((d) =>
      d && d.width === known.width && d.height === known.height &&
      Math.abs(deviceFps(d.frameRate) - known.fps) < FPS_TOLERANCE &&
      isInterlacedName(d.name) === !!known.interlaced);
    if (!found) continue;
    const offered = (found.videoModes || []).map((n) => DEVICE_PIXEL_FORMATS[n]).filter(Boolean);
    const pixelFormats = PIXEL_FORMAT_PREFERENCE.filter((f) => offered.includes(f));
    if (pixelFormats.length === 0) continue; // nothing the pipeline can send
    result.push({
      key, label: known.label, width: known.width, height: known.height,
      fps: known.fps, interlaced: !!known.interlaced, pixelFormats,
    });
  }
  return result;
}

module.exports = {
  DISPLAY_MODES, PIXEL_FORMAT_PREFERENCE, choosePixelFormat, filterOutputDevices, matchDeviceModes,
};
