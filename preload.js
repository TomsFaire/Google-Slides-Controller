const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  getDisplays: () => ipcRenderer.invoke('get-displays'),
  googleSignIn: () => ipcRenderer.invoke('google-signin'),
  googleSignOut: () => ipcRenderer.invoke('google-signout'),
  checkSignInStatus: () => ipcRenderer.invoke('check-signin-status'),
  openTestPresentation: () => ipcRenderer.invoke('open-test-presentation'),
  openPresentation: (data) => ipcRenderer.invoke('open-presentation', data),
  getPreferences: () => ipcRenderer.invoke('get-preferences'),
  savePreferences: (prefs) => ipcRenderer.invoke('save-preferences', prefs),
  getNetworkInfo: () => ipcRenderer.invoke('get-network-info'),
  getBuildInfo: () => ipcRenderer.invoke('get-build-info'),

  // Debug logs (desktop UI)
  getLogBuffer: () => ipcRenderer.invoke('get-log-buffer'),
  clearLogBuffer: () => ipcRenderer.invoke('clear-log-buffer'),
  exportLogBuffer: () => ipcRenderer.invoke('export-log-buffer'),
  onLogLine: (callback) => {
    if (typeof callback !== 'function') return;
    ipcRenderer.on('app-log-line', (_event, line) => callback(line));
  }
});
