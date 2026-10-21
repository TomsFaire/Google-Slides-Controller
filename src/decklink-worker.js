'use strict';

/**
 * decklink-worker.js
 * Runs in a child_process.fork() so that native DeckLink SDK crashes cannot
 * kill the Electron main process.  All messages use { id, cmd, ...payload }.
 * Every DeckLink call goes through the backend (see decklink-backend.js).
 */

const { createBackend } = require('./decklink-backend');

// `send(id, fields)` delivers a reply; injectable so the handler is testable.
function createHandler(backend, send) {
  const outputs = {}; // playbackId → { output, busy }

  return async function handle(msg) {
    const { id, cmd } = msg;

    switch (cmd) {
      case 'probe': {
        const { driverVersion } = await backend.probe();
        send(id, { cmd: 'ok', driverVersion });
        break;
      }

      case 'get_devices': {
        send(id, { cmd: 'ok', devices: await backend.listDevices() });
        break;
      }

      case 'start': {
        const output = await backend.openOutput({
          deviceIndex: msg.deviceIndex,
          mode: msg.bmdMode,
          width: msg.width,
          height: msg.height,
          pixelFormat: msg.pixelFormat || 'bgra',
        });
        outputs[msg.playbackId] = { output, busy: false };
        send(id, { cmd: 'ok' });
        break;
      }

      case 'frame': {
        const entry = outputs[msg.playbackId];
        if (!entry || entry.busy) {
          send(id, { cmd: 'skip' });
          break;
        }
        entry.busy = true;
        try {
          const timing = await entry.output.displayFrame(Buffer.from(msg.data));
          // Timings let the main process's stats line show which stage is slow.
          send(id, { cmd: 'ok', timing });
        } catch (e) {
          console.error('[decklink-worker] frame error:', e.message);
          throw e;
        } finally {
          entry.busy = false;
        }
        break;
      }

      case 'stop': {
        const entry = outputs[msg.playbackId];
        if (entry) {
          delete outputs[msg.playbackId];
          await entry.output.close().catch(() => {});
        }
        send(id, { cmd: 'ok' });
        break;
      }

      default:
        send(id, { cmd: 'error', error: `unknown command: ${cmd}` });
    }
  };
}

if (require.main === module) {
  const handle = createHandler(createBackend(), (id, fields) => process.send({ id, ...fields }));
  process.on('message', (msg) => {
    handle(msg).catch((e) => {
      try { process.send({ id: msg.id, cmd: 'error', error: e.message }); } catch (_) {}
    });
  });
}

module.exports = { createHandler };
