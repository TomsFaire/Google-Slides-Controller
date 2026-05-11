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

test('0x0c byte dispatches next-slide', () =>
  withServer(18899, {}, async (dispatched) => {
    await sendAndClose(18899, Buffer.from([0x0c]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide']);
  })
);

test('0x08 byte dispatches previous-slide', () =>
  withServer(18900, {}, async (dispatched) => {
    await sendAndClose(18900, Buffer.from([0x08]));
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
    await sendAndClose(18902, Buffer.from([0x0c, 0x08]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide', 'previous-slide']);
  })
);

test('keepalive bytes between commands do not affect dispatch', () =>
  withServer(18903, {}, async (dispatched) => {
    await sendAndClose(18903, Buffer.from([0xFF, 0x0c, 0xFF]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide']);
  })
);

test('onStatus transitions connected then listening on disconnect', () =>
  withServer(18904, {}, async (dispatched, statuses) => {
    await sendAndClose(18904, Buffer.from([0x0c]));
    await new Promise(r => setTimeout(r, 80));
    assert.ok(statuses.includes('connected'), `expected connected in [${statuses}]`);
    assert.ok(statuses.includes('listening'), `expected listening in [${statuses}]`);
  })
);

test('config.enabled = false suppresses dispatch', () =>
  withServer(18905, { config: { port: 18905, name: '', enabled: false } }, async (dispatched) => {
    await sendAndClose(18905, Buffer.from([0x0c]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, []);
  })
);

test('config.enabled = true allows dispatch', () =>
  withServer(18906, { config: { port: 18906, name: '', enabled: true } }, async (dispatched) => {
    await sendAndClose(18906, Buffer.from([0x0c]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide']);
  })
);

test('masterEnabled = () => false suppresses dispatch', () =>
  withServer(18907, { masterEnabled: () => false, config: { port: 18907, name: '', enabled: true } }, async (dispatched) => {
    await sendAndClose(18907, Buffer.from([0x0c]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, []);
  })
);

test('disabled port keeps TCP connection open (no dispatch)', () =>
  withServer(18908, { config: { port: 18908, name: '', enabled: false } }, async (dispatched, statuses) => {
    await sendAndClose(18908, Buffer.from([0x0c, 0x08]));
    await new Promise(r => setTimeout(r, 80));
    assert.deepEqual(dispatched, [], 'no commands dispatched when disabled');
    assert.ok(statuses.includes('connected'), 'connection was accepted even though disabled');
  })
);

test('waveshare adapter still dispatches commands', () =>
  withServer(18909, { config: { port: 18909, name: '', enabled: true, adapter: 'waveshare' } }, async (dispatched) => {
    await sendAndClose(18909, Buffer.from([0x0c]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide']);
  })
);

test('high-bit RS485 noise variant 0x8c dispatches next-slide', () =>
  withServer(18910, {}, async (dispatched) => {
    await sendAndClose(18910, Buffer.from([0x8c]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['next-slide']);
  })
);

test('high-bit RS485 noise variant 0x88 dispatches previous-slide', () =>
  withServer(18911, {}, async (dispatched) => {
    await sendAndClose(18911, Buffer.from([0x88]));
    await new Promise(r => setTimeout(r, 50));
    assert.deepEqual(dispatched, ['previous-slide']);
  })
);
