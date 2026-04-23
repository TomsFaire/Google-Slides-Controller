const net = require('net');
const { parsePerfectCueCommand, splitOnCR } = require('./perfectcue-parser');

function createPerfectCueServer({ dispatch, log, onStatus }) {
  const server = net.createServer(socket => {
    onStatus('connected', socket.remoteAddress);
    log(`DSAN connected from ${socket.remoteAddress}`);

    let buf = '';

    socket.on('data', chunk => {
      const hex = [...chunk].map(b => b.toString(16).padStart(2, '0')).join(' ');
      const ascii = chunk.toString('ascii').replace(/[^\x20-\x7e]/g, '.');
      log(`raw: ${hex} | ${ascii}`);

      buf += chunk.toString('ascii');
      const { lines, remainder } = splitOnCR(buf);
      buf = remainder;

      for (const line of lines) {
        log(`line: "${line}"`);
        const cmd = parsePerfectCueCommand(line);
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
