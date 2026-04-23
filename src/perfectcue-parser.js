function parsePerfectCueCommand(line) {
  if (line.includes('>FORWARD')) return 'next';
  if (line.includes('>REVERSE')) return 'previous';
  return null;
}

// Splits accumulated buffer on \r, returns completed lines and the leftover fragment.
// Call with the full buffer each time; replace buffer with remainder.
function splitOnCR(accumulated) {
  const parts = accumulated.split('\r');
  return {
    lines: parts.slice(0, -1).filter(Boolean),
    remainder: parts[parts.length - 1]
  };
}

module.exports = { parsePerfectCueCommand, splitOnCR };
