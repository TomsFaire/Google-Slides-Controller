const COMMAND_BYTES = {
  0x0F: 'next',
  0x1F: 'previous',
};

function parsePerfectCueByte(byte) {
  return COMMAND_BYTES[byte] ?? null;
}

module.exports = { parsePerfectCueByte };
