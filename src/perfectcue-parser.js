// WaveShare RS485 bridge occasionally sets the high bit due to framing noise; mask it off before lookup.
const COMMAND_BYTES = {
  0x0c: 'next',
  0x08: 'previous',
  0x04: 'blackout',
};

function parsePerfectCueByte(byte) {
  return COMMAND_BYTES[byte & 0x7f] ?? null;
}

module.exports = { parsePerfectCueByte };
