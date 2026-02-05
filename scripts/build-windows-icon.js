#!/usr/bin/env node
/**
 * Generate build/icon.ico from build/icon.png for Windows builds.
 * Run: node scripts/build-windows-icon.js
 * Requires: npm install --save-dev png-to-ico
 */
const path = require('path');
const fs = require('fs');

const pngPath = path.join(__dirname, '..', 'build', 'icon.png');
const icoPath = path.join(__dirname, '..', 'build', 'icon.ico');

if (!fs.existsSync(pngPath)) {
  console.error('Error: build/icon.png not found');
  process.exit(1);
}

async function main() {
  const pngToIco = require('png-to-ico');
  const buf = await pngToIco(pngPath);
  fs.writeFileSync(icoPath, buf);
  console.log('Created build/icon.ico');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
