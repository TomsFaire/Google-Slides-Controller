'use strict';

/**
 * macadam-worker.js
 * Runs in a child_process.fork() so that native DeckLink SDK crashes cannot
 * kill the Electron main process.  All messages use { id, cmd, ...payload }.
 */

let macadam = null;

function reply(id, fields) {
  process.send({ id, ...fields });
}

async function handle(msg) {
  const { id, cmd } = msg;

  switch (cmd) {
    // ── probe ─────────────────────────────────────────────────────────────
    // Just verify the native addon loads and exposes constants.
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

    default:
      reply(id, { cmd: 'error', error: `unknown command: ${cmd}` });
  }
}

process.on('message', (msg) => {
  handle(msg).catch((e) => {
    try { reply(msg.id, { cmd: 'error', error: e.message }); } catch (_) {}
  });
});
