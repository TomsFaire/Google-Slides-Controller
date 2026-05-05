const net = require('net');
const { parsePerfectCueByte } = require('./perfectcue-parser');

function createPerfectCueServer({ config = null, masterEnabled = null, isAllowed = null, dispatch, log, onStatus }) {
  const server = net.createServer(socket => {
    const remoteIp = socket.remoteAddress;
    if (typeof isAllowed === 'function' && !isAllowed(remoteIp)) {
      log(`connection from ${remoteIp} rejected (not in allowlist)`);
      socket.destroy();
      return;
    }
    onStatus('connected', remoteIp);
    log(`DSAN connected from ${remoteIp}`);
    // Use OS default timing for probes (short initial delays upset some DSAN / PerfectCue links after idle periods).
    socket.setKeepAlive(true, 0);

    // Send 0xFF every 15 s so the USR-TCP232's idle timer never fires.
    // 0xFF is a recognised no-op by parsePerfectCueByte and is ignored by the PerfectCue receiver.
    const PING_INTERVAL_MS = 15_000;
    const pingTimer = setInterval(() => {
      if (!socket.destroyed) socket.write(Buffer.from([0xff]));
    }, PING_INTERVAL_MS);

    // If no data arrives for 50 s (ping failed to get through → dead connection),
    // destroy our side so the USR-TCP232 sees a RST and reconnects.
    const IDLE_TIMEOUT_MS = 50_000;
    socket.setTimeout(IDLE_TIMEOUT_MS);
    socket.on('timeout', () => {
      log('idle timeout — closing socket to force reconnect');
      socket.destroy();
    });

    socket.on('data', chunk => {
      const hex = [...chunk].map(b => b.toString(16).padStart(2, '0')).join(' ');
      const ascii = chunk.toString('ascii').replace(/[^\x20-\x7e]/g, '.');
      log(`raw: ${hex} | ${ascii}`);

      for (const byte of chunk) {
        const cmd = parsePerfectCueByte(byte);
        if (cmd !== 'next' && cmd !== 'previous') continue;
        const globalEnabled = masterEnabled === null || masterEnabled() === true;
        const portEnabled = config === null || config.enabled !== false;
        if (globalEnabled && portEnabled) {
          dispatch(cmd === 'next' ? 'next-slide' : 'previous-slide');
        }
      }
    });

    socket.on('close', () => {
      clearInterval(pingTimer);
      onStatus('listening', null);
      log('DSAN disconnected, waiting for reconnect');
    });

    socket.on('error', err => {
      clearInterval(pingTimer);
      log(`socket error: ${err.message}`);
    });
  });

  server.on('error', err => {
    onStatus('error', null);
    log(`port error: ${err.message}`);
  });

  return server;
}

module.exports = { createPerfectCueServer };
