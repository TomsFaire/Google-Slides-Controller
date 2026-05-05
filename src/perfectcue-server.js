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
      onStatus('listening', null);
      log('DSAN disconnected, waiting for reconnect');
    });

    socket.on('error', err => {
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
