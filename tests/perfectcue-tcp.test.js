const { test } = require('node:test');
const assert = require('node:assert/strict');
const net = require('net');
const { createPerfectCueServer } = require('../src/perfectcue-server');

function sendAndClose(port, data) {
  return new Promise((resolve, reject) => {
    const client = net.createConnection(port, '127.0.0.1', () => {
      client.write(data);
      setTimeout(() => { client.destroy(); resolve(); }, 30);
    });
    client.on('error', reject);
  });
}

function withServer(port, overrides, fn) {
  const dispatched = [];
  const statuses = [];
  const server = createPerfectCueServer({
    dispatch: (endpoint) => dispatched.push(endpoint),
    log: () => {},
    onStatus: (status) => statuses.push(status),
    ...overrides
  });
  return new Promise((resolve, reject) => {
    server.listen(port, '127.0.0.1', async () => {
      try {
        await fn(dispatched, statuses);
        resolve();
      } catch (e) {
        reject(e);
      } finally {
        server.close();
      }
    });
    server.on('error', reject);
  });
}

test('>FORWARD CR dispatches next-slide', () =>
  withServer(18899, {}, async (dispatched) => {
    await sendAndClose(18899, '>FORWARD 35\r');
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide']);
  })
);

test('>REVERSE CR dispatches previous-slide', () =>
  withServer(18900, {}, async (dispatched) => {
    await sendAndClose(18900, '>REVERSE 3C\r');
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['previous-slide']);
  })
);

test('unknown command does not dispatch', () =>
  withServer(18901, {}, async (dispatched) => {
    await sendAndClose(18901, '>STATUS OK\r');
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, []);
  })
);

test('multiple commands in one chunk all dispatch', () =>
  withServer(18902, {}, async (dispatched) => {
    await sendAndClose(18902, '>FORWARD 35\r>REVERSE 3C\r');
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide', 'previous-slide']);
  })
);

test('command split across two chunks dispatches once', () =>
  withServer(18903, {}, async (dispatched) => {
    await new Promise((resolve, reject) => {
      const client = net.createConnection(18903, '127.0.0.1', () => {
        client.write('>FORW');
        setTimeout(() => {
          client.write('ARD 35\r');
          setTimeout(() => { client.destroy(); resolve(); }, 30);
        }, 20);
      });
      client.on('error', reject);
    });
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide']);
  })
);

test('onStatus transitions connected then listening on disconnect', () =>
  withServer(18904, {}, async (dispatched, statuses) => {
    await sendAndClose(18904, '>FORWARD 35\r');
    await new Promise(r => setTimeout(r, 80));
    assert.ok(statuses.includes('connected'), `expected connected in [${statuses}]`);
    assert.ok(statuses.includes('listening'), `expected listening in [${statuses}]`);
  })
);
