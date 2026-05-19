'use strict';

/**
 * macadam-worker.js
 * Runs in a child_process.fork() so that native DeckLink SDK crashes cannot
 * kill the Electron main process.  All messages use { id, cmd, ...payload }.
 */

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
      console.error(`[decklink-worker] start: deviceIndex=${msg.deviceIndex} bmdMode=${msg.bmdMode} modeVal=${macadam[msg.bmdMode]} pixelFormat=${macadam.bmdFormat8BitBGRA}`);
      const pb = await macadam.playback({
        deviceIndex: msg.deviceIndex,
        displayMode: macadam[msg.bmdMode],
        pixelFormat: macadam.bmdFormat8BitBGRA,
      });
      playbacks[msg.playbackId] = { pb, busy: false };
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
      const buf = Buffer.from(msg.data);
      await new Promise((resolve) => {
        entry.pb.frame(buf, () => {
          entry.busy = false;
          resolve();
        });
      });
      reply(id, { cmd: 'ok' });
      break;
    }

    // ── stop ──────────────────────────────────────────────────────────────
    case 'stop': {
      const entry = playbacks[msg.playbackId];
      if (entry) {
        try { entry.pb.stop(); } catch (_) {}
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
