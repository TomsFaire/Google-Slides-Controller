'use strict';

/**
 * Dumps what a real DeckLink device reports about itself: attributes plus the
 * output display modes (and pixel formats) the driver says it supports.
 * Read-only: does not start any output.
 *
 *   yarn decklink:info
 *   DECKLINK_DEVICE=1 yarn decklink:info
 *
 * Runs under ELECTRON_RUN_AS_NODE (see package.json) because macadam is
 * rebuilt for Electron.
 */

const macadam = require('macadam');

const DEVICE_INDEX = Number(process.env.DECKLINK_DEVICE || 0);

async function main() {
  const devices = await macadam.getDeviceInfo();
  console.log(`${devices.length} device(s) found`);
  const d = devices[DEVICE_INDEX];
  if (!d) {
    console.error(`no device at index ${DEVICE_INDEX}`);
    process.exit(1);
  }

  const { outputDisplayModes, inputDisplayModes, ...attrs } = d;
  console.log(`\n== device ${DEVICE_INDEX} attributes ==`);
  console.log(JSON.stringify(attrs, null, 2));

  console.log(`\n== output display modes (${(outputDisplayModes || []).length}) ==`);
  for (const m of outputDisplayModes || []) console.log(JSON.stringify(m));
}

main().catch((e) => { console.error(e); process.exit(2); });
