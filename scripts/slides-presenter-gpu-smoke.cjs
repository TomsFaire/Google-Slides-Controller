/**
 * Minimal Electron window loading a Google Slides presenter URL.
 * Compare rendering vs Chrome and vs the full app (same partition optional).
 *
 * Usage:
 *   yarn smoke:slides-gpu "https://docs.google.com/presentation/d/.../present"
 *
 * Optional env:
 *   GSLIDE_SESSION_PARTITION=persist:google   # reuse app Google session (optional)
 */
const { app, BrowserWindow } = require('electron');

const url = process.argv[2];
if (!url || !/^https?:\/\//i.test(url)) {
  console.error('Usage: yarn smoke:slides-gpu "<presenter https URL>"');
  process.exit(1);
}

const partition = (process.env.GSLIDE_SESSION_PARTITION || '').trim();

app.whenReady().then(() => {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      ...(partition ? { partition } : {})
    }
  });
  win.loadURL(url);
  if (process.platform === 'darwin') {
    try {
      win.setSimpleFullScreen(true);
    } catch (e) {
      // ignore
    }
  }
});

app.on('window-all-closed', () => {
  app.quit();
});
