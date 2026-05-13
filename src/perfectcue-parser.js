// RS485 framing noise commonly corrupts the two high bits (bits 7+6); mask both before lookup.
const COMMAND_BYTES = {
  0x0c: 'next',
  0x06: 'next',   // 0x0c right-shifted 1 bit (RS485 start-bit mis-frame)
  0x08: 'previous',
  0x04: 'blackout',
};

function parsePerfectCueByte(byte) {
  return COMMAND_BYTES[byte & 0x3f] ?? null;
}

module.exports = { parsePerfectCueByte };
