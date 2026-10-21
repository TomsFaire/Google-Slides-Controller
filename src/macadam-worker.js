'use strict';

/**
 * macadam-worker.js
 * Runs in a child_process.fork() so that native DeckLink SDK crashes cannot
 * kill the Electron main process.  All messages use { id, cmd, ...payload }.
 */

const { bgraToUyvy } = require('./bgra-to-uyvy');

// Defaults (709 / limited) are right for HD monitors. The env vars exist only
// to A/B colour problems: DECKLINK_YUV_MATRIX=709|601 DECKLINK_YUV_RANGE=limited|full
const YUV_OPTIONS = {
  matrix: process.env.DECKLINK_YUV_MATRIX || '709',
  range: process.env.DECKLINK_YUV_RANGE || 'limited',
};

let macadam = null;
const playbacks = {}; // playbackId → { pb, busy }

function reply(id, fields) {
  process.send({ id, ...fields });
}

async function handle(msg) {
  const { id, cmd } = msg;

  switch (cmd) {
    // ── probe ─────────────────────────────────────────────────────────────
    case 'probe': {
      macadam = require('macadam');
      if (macadam.bmdModeHD1080p2997 === undefined)
        throw new Error('macadam constants missing — rebuild required');
      reply(id, { cmd: 'ok' });
      break;
    }

    // ── get_devices ───────────────────────────────────────────────────────
    case 'get_devices': {
      if (!macadam) macadam = require('macadam');
      const info = await macadam.getDeviceInfo();
      reply(id, {
        cmd: 'ok',
        devices: info.map((d, i) => ({
          index: i,
          name: d.displayName || d.modelName || `DeckLink ${i}`,
        })),
      });
      break;
    }

    // ── start ─────────────────────────────────────────────────────────────
    case 'start': {
      if (!macadam) macadam = require('macadam');
      const yuv = msg.pixelFormat === 'yuv';
      const pixelFormat = yuv ? macadam.bmdFormat8BitYUV : macadam.bmdFormat8BitBGRA;
      console.error(`[decklink-worker] start: deviceIndex=${msg.deviceIndex} bmdMode=${msg.bmdMode} modeVal=${macadam[msg.bmdMode]} pixelFormat=${msg.pixelFormat || 'bgra'}(${pixelFormat})`);
      const pb = await macadam.playback({
        deviceIndex: msg.deviceIndex,
        displayMode: macadam[msg.bmdMode],
        pixelFormat,
      });
      playbacks[msg.playbackId] = {
        pb,
        busy: false,
        yuv,
        width: msg.width,
        height: msg.height,
        yuvBuf: yuv ? Buffer.alloc(msg.width * msg.height * 2) : null,
      };
      reply(id, { cmd: 'ok' });
      break;
    }

    // ── frame ─────────────────────────────────────────────────────────────
    case 'frame': {
      const entry = playbacks[msg.playbackId];
      if (!entry || entry.busy) {
        reply(id, { cmd: 'skip' });
        break;
      }
      entry.busy = true;
      const t0  = Date.now();
      const raw = Buffer.from(msg.data);
      const buf = entry.yuv
        ? bgraToUyvy(raw, entry.width, entry.height, entry.yuvBuf, YUV_OPTIONS)
        : raw;
      const t1  = Date.now();
      let timer;
      try {
        // macadam's native playback object: displayFrame() resolves once the
        // frame has been handed to the hardware (it has no .frame()/.on()).
        await Promise.race([
          entry.pb.displayFrame(buf),
          new Promise((_, reject) => {
            timer = setTimeout(
              () => reject(new Error('displayFrame() timed out after 2000ms')), 2000);
          }),
        ]);
      } catch (e) {
        console.error('[decklink-worker] frame error:', e.message);
        throw e;
      } finally {
        clearTimeout(timer);
        entry.busy = false;
      }
      // Timings let the main process's stats line show which stage is slow.
      reply(id, { cmd: 'ok', timing: { convertMs: t1 - t0, displayMs: Date.now() - t1 } });
      break;
    }

    // ── stop ──────────────────────────────────────────────────────────────
    case 'stop': {
      const entry = playbacks[msg.playbackId];
      if (entry) {
        try { Promise.resolve(entry.pb.stop()).catch(() => {}); } catch (_) {}
        delete playbacks[msg.playbackId];
      }
      reply(id, { cmd: 'ok' });
      break;
    }

    default:
      reply(id, { cmd: 'error', error: `unknown command: ${cmd}` });
  }
}

process.on('message', (msg) => {
  handle(msg).catch((e) => {
    try { reply(msg.id, { cmd: 'error', error: e.message }); } catch (_) {}
  });
});
