const net = require('net');
const { parsePerfectCueByte } = require('./perfectcue-parser');

function createPerfectCueServer({ dispatch, isAllowed, log, onStatus }) {
  const server = net.createServer(socket => {
    const remoteIp = socket.remoteAddress;
    if (isAllowed && !isAllowed(remoteIp)) {
      log(`connection from ${remoteIp} rejected (not in allowlist)`);
      socket.destroy();
      return;
    }
    onStatus('connected', remoteIp);
    log(`DSAN connected from ${remoteIp}`);

    socket.on('data', chunk => {
      const hex = [...chunk].map(b => b.toString(16).padStart(2, '0')).join(' ');
      const ascii = chunk.toString('ascii').replace(/[^\x20-\x7e]/g, '.');
      log(`raw: ${hex} | ${ascii}`);

      for (const byte of chunk) {
        const cmd = parsePerfectCueByte(byte);
        if (cmd === 'next') dispatch('next-slide');
        else if (cmd === 'previous') dispatch('previous-slide');
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
