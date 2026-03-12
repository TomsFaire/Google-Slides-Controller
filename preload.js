const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  getDisplays: () => ipcRenderer.invoke('get-displays'),
  googleSignIn: () => ipcRenderer.invoke('google-signin'),
  googleSignOut: () => ipcRenderer.invoke('google-signout'),
  checkSignInStatus: () => ipcRenderer.invoke('check-signin-status'),
  openTestPresentation: () => ipcRenderer.invoke('open-test-presentation'),
  openPresentation: (data) => ipcRenderer.invoke('open-presentation', data),
  openPresentationEdit: () => ipcRenderer.invoke('open-presentation-edit'),
  switchPresentationToPresent: () => ipcRenderer.invoke('switch-presentation-to-present'),
  getPreferences: () => ipcRenderer.invoke('get-preferences'),
  getSpeakerNotes: () => ipcRenderer.invoke('get-speaker-notes'),
  savePreferences: (prefs) => ipcRenderer.invoke('save-preferences', prefs),
  showOpenCssDialog: () => ipcRenderer.invoke('show-open-css-dialog'),
  showOpenLogoDialog: () => ipcRenderer.invoke('show-open-logo-dialog'),
  showOpenCertDialog: () => ipcRenderer.invoke('show-open-cert-dialog'),
  showOpenKeyDialog: () => ipcRenderer.invoke('show-open-key-dialog'),
  showOpenSlidoExtensionDialog: () => ipcRenderer.invoke('show-open-slido-extension-dialog'),
  downloadCssTemplate: () => ipcRenderer.invoke('download-css-template'),
  getNetworkInfo: () => ipcRenderer.invoke('get-network-info'),
  getBuildInfo: () => ipcRenderer.invoke('get-build-info'),
  getLastPresentationUrl: () => ipcRenderer.invoke('get-last-presentation-url'),
  getPresentationEditUrl: () => ipcRenderer.invoke('get-presentation-edit-url'),
  openUrlInBrowser: (url) => ipcRenderer.invoke('open-url-in-browser', url),

  // Debug logs (desktop UI)
  getLogBuffer: () => ipcRenderer.invoke('get-log-buffer'),
  clearLogBuffer: () => ipcRenderer.invoke('clear-log-buffer'),
  exportLogBuffer: () => ipcRenderer.invoke('export-log-buffer'),
  getCrashInfo: () => ipcRenderer.invoke('get-crash-info'),
  openCrashReportsFolder: () => ipcRenderer.invoke('open-crash-reports-folder'),
  onLogLine: (callback) => {
    if (typeof callback !== 'function') return;
    ipcRenderer.on('app-log-line', (_event, line) => callback(line));
  }
});
