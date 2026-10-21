'use strict';

/**
 * decklink-worker.js
 * Runs in a child_process.fork() so that native DeckLink SDK crashes cannot
 * kill the Electron main process.  All messages use { id, cmd, ...payload }.
 * Every DeckLink call goes through the backend (see decklink-backend.js).
 */

const { createBackend, openOutputNegotiated } = require('./decklink-backend');

// The device already says which pixel formats it takes for each mode (a query
// that does not open the output), so use that and open once. Opening in a format
// the device rejects first is avoided: only if the mode list is unavailable do
// we fall back to trying each format in turn.
async function formatsFor(backend, msg) {
  if (msg.pixelFormats) return msg.pixelFormats;
  try {
    const modes = await backend.listModes(msg.deviceIndex);
    const mode = modes.find((m) => m.bmdMode === msg.bmdMode);
    if (mode && mode.pixelFormats.length) return mode.pixelFormats;
  } catch (_) { /* fall through to trial */ }
  return undefined;
}

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

      case 'get_modes': {
        send(id, { cmd: 'ok', modes: await backend.listModes(msg.deviceIndex) });
        break;
      }

      case 'start': {
        const { output, pixelFormat } = await openOutputNegotiated(backend, {
          deviceIndex: msg.deviceIndex,
          mode: msg.bmdMode,
          width: msg.width,
          height: msg.height,
        }, await formatsFor(backend, msg));
        outputs[msg.playbackId] = { output, busy: false };
        send(id, { cmd: 'ok', pixelFormat });
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
