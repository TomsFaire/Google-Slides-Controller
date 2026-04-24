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

test('0x0F byte dispatches next-slide', () =>
  withServer(18899, {}, async (dispatched) => {
    await sendAndClose(18899, Buffer.from([0x0F]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide']);
  })
);

test('0x1F byte dispatches previous-slide', () =>
  withServer(18900, {}, async (dispatched) => {
    await sendAndClose(18900, Buffer.from([0x1F]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['previous-slide']);
  })
);

test('0xFF keepalive does not dispatch', () =>
  withServer(18901, {}, async (dispatched) => {
    await sendAndClose(18901, Buffer.from([0xFF]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, []);
  })
);

test('multiple bytes in one chunk all dispatch', () =>
  withServer(18902, {}, async (dispatched) => {
    await sendAndClose(18902, Buffer.from([0x0F, 0x1F]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide', 'previous-slide']);
  })
);

test('keepalive bytes between commands do not affect dispatch', () =>
  withServer(18903, {}, async (dispatched) => {
    await sendAndClose(18903, Buffer.from([0xFF, 0x0F, 0xFF]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide']);
  })
);

test('onStatus transitions connected then listening on disconnect', () =>
  withServer(18904, {}, async (dispatched, statuses) => {
    await sendAndClose(18904, Buffer.from([0x0F]));
    await new Promise(r => setTimeout(r, 80));
    assert.ok(statuses.includes('connected'), `expected connected in [${statuses}]`);
    assert.ok(statuses.includes('listening'), `expected listening in [${statuses}]`);
  })
);
