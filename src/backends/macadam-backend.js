'use strict';

const { bgraToUyvy } = require('../bgra-to-uyvy');
const { matchDeviceModes } = require('../decklink-modes');

const DISPLAY_FRAME_TIMEOUT_MS = 2000;

// Defaults (709 / limited) are right for HD monitors. The env vars exist only
// to A/B colour problems: DECKLINK_YUV_MATRIX=709|601 DECKLINK_YUV_RANGE=limited|full
function yuvOptions() {
  return {
    matrix: process.env.DECKLINK_YUV_MATRIX || '709',
    range: process.env.DECKLINK_YUV_RANGE || 'limited',
  };
}

// `load` returns the macadam module; injectable so tests need no hardware.
function createMacadamBackend(load = () => require('macadam')) {
  let macadam = null;
  const mod = () => (macadam = macadam || load());

  return {
    async probe() {
      const m = mod();
      if (m.bmdModeHD1080p2997 === undefined)
        throw new Error('macadam constants missing — rebuild required');
      // "DeckLinkAPI version: 16.4.0". Throws when Desktop Video is not installed.
      let driverVersion = null;
      try { driverVersion = m.deckLinkVersion().replace(/^DeckLinkAPI version:\s*/, 'DeckLink API '); }
      catch (_) { /* no driver: device listing will report it */ }
      return { driverVersion };
    },

    async listDevices() {
      const info = await mod().getDeviceInfo();
      // macadam does not say whether a device can output, so assume it can.
      return info.map((d, index) => ({
        index,
        name: d.displayName || d.modelName || `DeckLink ${index}`,
        canOutput: true,
      }));
    },

    // Modes the device lists for output that the pipeline can drive.
    async listModes(deviceIndex) {
      const info = await mod().getDeviceInfo();
      const device = info[deviceIndex];
      return device ? matchDeviceModes(device.outputDisplayModes) : [];
    },

    async openOutput({ deviceIndex, mode, width, height, pixelFormat }) {
      const m = mod();
      const yuv = pixelFormat === 'yuv';
      console.error(`[decklink-worker] start: deviceIndex=${deviceIndex} mode=${mode} pixelFormat=${pixelFormat}`);
      const pb = await m.playback({
        deviceIndex,
        displayMode: m[mode],
        pixelFormat: yuv ? m.bmdFormat8BitYUV : m.bmdFormat8BitBGRA,
      });
      const yuvBuf = yuv ? Buffer.alloc(width * height * 2) : null;
      const options = yuvOptions();

      return {
        async displayFrame(bgra) {
          const t0 = Date.now();
          const buf = yuv ? bgraToUyvy(bgra, width, height, yuvBuf, options) : bgra;
          const t1 = Date.now();
          let timer;
          try {
            // displayFrame() resolves once the frame has been handed to the hardware.
            await Promise.race([
              pb.displayFrame(buf),
              new Promise((_, reject) => {
                timer = setTimeout(() => reject(new Error(
                  `displayFrame() timed out after ${DISPLAY_FRAME_TIMEOUT_MS}ms`)), DISPLAY_FRAME_TIMEOUT_MS);
              }),
            ]);
          } finally {
            clearTimeout(timer);
          }
          return { convertMs: t1 - t0, displayMs: Date.now() - t1 };
        },
        async close() {
          await Promise.resolve(pb.stop()).catch(() => {});
        },
      };
    },
  };
}

module.exports = { createMacadamBackend };
