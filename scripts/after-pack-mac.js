/**
 * Electron-builder afterPack hook for macOS.
 * Clears quarantine extended attributes from the built .app so that when the app
 * is zipped and distributed, testers can open it without macOS reporting it as "damaged"
 * (a common issue with unsigned apps). After extracting the zip, recipients may still
 * need to run: xattr -cr "Google Slides Opener.app" if the zip was downloaded.
 */
const path = require('path');
const { execSync } = require('child_process');

module.exports = async function (context) {
  if (context.electronPlatformName !== 'darwin') return;
  const appName = context.packager.appInfo.productFilename;
  const appPath = path.join(context.appOutDir, `${appName}.app`);
  try {
    execSync(`xattr -cr "${appPath}"`, { stdio: 'inherit' });
  } catch (e) {
    console.warn('[after-pack-mac] xattr -cr failed (non-fatal):', e.message);
  }
};
