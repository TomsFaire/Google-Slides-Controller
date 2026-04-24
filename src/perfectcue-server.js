const net = require('net');
const { parsePerfectCueByte } = require('./perfectcue-parser');

function createPerfectCueServer({ config = null, masterEnabled = null, dispatch, log, onStatus }) {
  const server = net.createServer(socket => {
    onStatus('connected', socket.remoteAddress);
    log(`DSAN connected from ${socket.remoteAddress}`);

    socket.on('data', chunk => {
      const hex = [...chunk].map(b => b.toString(16).padStart(2, '0')).join(' ');
      const ascii = chunk.toString('ascii').replace(/[^\x20-\x7e]/g, '.');
      log(`raw: ${hex} | ${ascii}`);

      for (const byte of chunk) {
        const cmd = parsePerfectCueByte(byte);
        if (cmd !== 'next' && cmd !== 'previous') continue;
        const globalEnabled = masterEnabled ? masterEnabled() : true;
        const portEnabled = config ? config.enabled !== false : true;
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
