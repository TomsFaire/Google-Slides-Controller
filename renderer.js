// DOM Elements
const presentationDisplay = document.getElementById('presentation-display');
const notesDisplay = document.getElementById('notes-display');
const machineNameInput = document.getElementById('machine-name');
const apiPortInput = document.getElementById('api-port');
const webUiPortInput = document.getElementById('web-ui-port');
const signinBtn = document.getElementById('signin-btn');
const testBtn = document.getElementById('test-btn');
const statusMessage = document.getElementById('status-message');
const modePrimary = document.getElementById('mode-primary');
const modeBackup = document.getElementById('mode-backup');
const modeStandalone = document.getElementById('mode-standalone');
const backupConfig = document.getElementById('backup-config');
const backupPortInput = document.getElementById('backup-port');
const backupIpList = document.getElementById('backup-ip-list');
const addBackupIpBtn = document.getElementById('add-backup-ip');
const presetList = document.getElementById('preset-list');
const addPresetBtn = document.getElementById('add-preset');
const savePresetsBtn = document.getElementById('save-presets-btn');
const loadPresetsBtn = document.getElementById('load-presets-btn');
const stagetimerRoomIdInput = document.getElementById('stagetimer-room-id');
const stagetimerApiKeyInput = document.getElementById('stagetimer-api-key');
const stagetimerEnabledCheckbox = document.getElementById('stagetimer-enabled');
const stagetimerVisibleCheckbox = document.getElementById('stagetimer-visible');
const saveStagetimerBtn = document.getElementById('save-stagetimer-btn');
const loadStagetimerBtn = document.getElementById('load-stagetimer-btn');
const perfectCueEnabledCheckbox = document.getElementById('perfectcue-enabled');
const perfectCuePortList = document.getElementById('perfectcue-port-list');
const addPerfectCuePortBtn = document.getElementById('add-perfectcue-port');
const savePerfectCueBtn = document.getElementById('save-perfectcue-btn');
const verboseLoggingCheckbox = document.getElementById('verbose-logging');
const webUiDebugConsoleEnabledCheckbox = document.getElementById('web-ui-debug-console-enabled');
const allowArbitraryUrlCheckbox = document.getElementById('allow-arbitrary-url');
const genericUrlBgColorPicker = document.getElementById('generic-url-background-color');
const genericUrlBgColorHex = document.getElementById('generic-url-background-color-hex');
const webUiThemeSelect = document.getElementById('web-ui-theme');
const webUiLogoPathInput = document.getElementById('web-ui-logo-path');
const webUiLogoChooseBtn = document.getElementById('web-ui-logo-choose');
const webUiLogoClearBtn = document.getElementById('web-ui-logo-clear');
const webUiCustomCssPathInput = document.getElementById('web-ui-custom-css-path');
const webUiCustomCssChooseBtn = document.getElementById('web-ui-custom-css-choose');
const webUiCustomCssClearBtn = document.getElementById('web-ui-custom-css-clear');
const saveWebUiAppearanceBtn = document.getElementById('save-web-ui-appearance-btn');
const keyboardShortcutPresetSelect = document.getElementById('keyboard-shortcut-preset');
const keyboardShortcutsDefaultEnabledCheckbox = document.getElementById('keyboard-shortcuts-default-enabled');
const saveKeyboardShortcutsBtn = document.getElementById('save-keyboard-shortcuts-btn');
const webUiDownloadCssTemplateBtn = document.getElementById('web-ui-download-css-template');
const webUiUseHttpsCheckbox = document.getElementById('web-ui-use-https');
const webUiCertPathInput = document.getElementById('web-ui-cert-path');
const webUiKeyPathInput = document.getElementById('web-ui-key-path');
const webUiCertChooseBtn = document.getElementById('web-ui-cert-choose');
const webUiCertClearBtn = document.getElementById('web-ui-cert-clear');
const webUiKeyChooseBtn = document.getElementById('web-ui-key-choose');
const webUiKeyClearBtn = document.getElementById('web-ui-key-clear');
const webUiHttpsCertGroup = document.getElementById('web-ui-https-cert-group');
const webUiHttpsKeyGroup = document.getElementById('web-ui-https-key-group');
const controllerIpList = document.getElementById('controller-ip-list');
const addControllerIpBtn = document.getElementById('add-controller-ip');
const speakerNotesCapture = document.getElementById('speaker-notes-capture');
const speakerNotesRefreshBtn = document.getElementById('speaker-notes-refresh');
const speakerNotesCopyBtn = document.getElementById('speaker-notes-copy');
const debugLogsConsole = document.getElementById('debug-logs-console');
const debugLogsClearBtn = document.getElementById('debug-logs-clear');
const debugLogsSaveBtn = document.getElementById('debug-logs-save');
const crashReportsDirEl = document.getElementById('crash-reports-dir');
const crashDumpsDirEl = document.getElementById('crash-dumps-dir');
const lastCrashTimeEl = document.getElementById('last-crash-time');
const openCrashReportsFolderBtn = document.getElementById('open-crash-reports-folder-btn');
const wanEnabledCheckbox = document.getElementById('wan-enabled');
const wanStatusRow = document.getElementById('wan-status-row');
const wanUrlDisplay = document.getElementById('wan-url-display');
const wanCopyBtn = document.getElementById('wan-copy-btn');
const wanTunnelPinNewInput = document.getElementById('wan-tunnel-pin-new');
const wanTunnelPinConfirmInput = document.getElementById('wan-tunnel-pin-confirm');
const wanTunnelPinSaveBtn = document.getElementById('wan-tunnel-pin-save');
const wanTunnelPinRemoveBtn = document.getElementById('wan-tunnel-pin-remove');
const wanTunnelPinStatusEl = document.getElementById('wan-tunnel-pin-status');
const wanPinScopeSelect = document.getElementById('wan-pin-scope');

// Tab switching
const NAV_ITEMS = document.querySelectorAll('.sidebar-nav [data-target]');
const TAB_SECTIONS = document.querySelectorAll('.content section[data-tab]');
const ACTIVE_TAB_KEY = 'desktop-ui:activeTab';

function switchTab(tabName) {
  TAB_SECTIONS.forEach(s => {
    s.hidden = s.dataset.tab !== tabName;
  });
  NAV_ITEMS.forEach(a => {
    a.classList.toggle('active', a.dataset.target === tabName);
  });
  try { localStorage.setItem(ACTIVE_TAB_KEY, tabName); } catch (_) {}
}

NAV_ITEMS.forEach(item => {
  item.addEventListener('click', () => switchTab(item.dataset.target));
  item.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); switchTab(item.dataset.target); }
  });
});

// Restore last tab on load
try {
  const saved = localStorage.getItem(ACTIVE_TAB_KEY);
  if (saved && document.querySelector(`section[data-tab="${saved}"]`)) {
    switchTab(saved);
  }
} catch (_) {}

// Status bar
function updateStatusBar() {
  const machineEl = document.getElementById('status-machine-name');
  if (machineEl && machineNameInput) {
    machineEl.textContent = machineNameInput.value.trim() || 'This machine';
  }

  // Sign-in LED
  const signinLed = document.getElementById('status-signin-led');
  const signinLabel = document.getElementById('status-signin-label');
  if (signinLed && signinLabel) {
    if (signinBtn && signinBtn.disabled) {
      signinLed.className = 'led led-warn';
      signinLabel.textContent = 'Signing in\u2026';
    } else if (isSignedIn) {
      signinLed.className = 'led led-ok';
      signinLabel.textContent = 'Signed in';
    } else {
      signinLed.className = 'led led-bad';
      signinLabel.textContent = 'Not signed in';
    }
  }

  // Port label
  const portLabel = document.getElementById('status-port-label');
  if (portLabel && apiPortInput) {
    portLabel.textContent = `:${apiPortInput.value || '9595'}`;
  }

  // Tunnel LED
  const tunnelLed = document.getElementById('status-tunnel-led');
  const tunnelLabel = document.getElementById('status-tunnel-label');
  if (tunnelLed && tunnelLabel) {
    const tunnelOn = wanEnabledCheckbox && wanEnabledCheckbox.checked;
    tunnelLed.className = tunnelOn ? 'led led-ok' : 'led led-neutral';
    tunnelLabel.textContent = tunnelOn ? 'Tunnel active' : 'LAN only';
  }

  // Mode label
  const modeLabel = document.getElementById('status-mode-label');
  if (modeLabel) {
    if (modePrimary && modePrimary.checked) {
      modeLabel.textContent = 'Primary';
    } else if (modeBackup && modeBackup.checked) {
      modeLabel.textContent = 'Backup';
    } else {
      modeLabel.textContent = 'Standalone';
    }
  }
}

// Dashboard WAN mirror
const dashboardWanEnabled = document.getElementById('dashboard-wan-enabled');
const dashboardWanUrlRow = document.getElementById('dashboard-wan-url-row');
const dashboardWanUrlDisplay = document.getElementById('dashboard-wan-url-display');

function syncDashboardWan() {
  if (!wanEnabledCheckbox) return;
  if (dashboardWanEnabled) dashboardWanEnabled.checked = wanEnabledCheckbox.checked;
  if (dashboardWanUrlRow) dashboardWanUrlRow.style.display = wanEnabledCheckbox.checked ? '' : 'none';
  if (dashboardWanUrlDisplay && wanUrlDisplay) {
    dashboardWanUrlDisplay.value = wanUrlDisplay.value;
  }
}

if (dashboardWanEnabled) {
  dashboardWanEnabled.addEventListener('change', () => {
    if (wanEnabledCheckbox) {
      wanEnabledCheckbox.checked = dashboardWanEnabled.checked;
      wanEnabledCheckbox.dispatchEvent(new Event('change'));
    }
  });
}

// Inline machine name rename on Dashboard
function initMachineNameRename() {
  const heading = document.getElementById('dashboard-machine-heading');
  const display = document.getElementById('dashboard-machine-display');
  if (!heading || !display || !machineNameInput) return;

  function refreshDisplay() {
    display.textContent = machineNameInput.value.trim() || 'This Machine';
  }
  refreshDisplay();

  function startEditing() {
    const current = machineNameInput.value.trim();
    const input = document.createElement('input');
    input.type = 'text';
    input.value = current;
    input.maxLength = 50;
    input.style.cssText = 'font: inherit; border: none; outline: 2px solid var(--text); border-radius: 2px; background: transparent; width: 100%; min-width: 120px; padding: 0 2px;';
    display.replaceWith(input);
    input.select();

    function commit() {
      machineNameInput.value = input.value.trim();
      machineNameInput.dispatchEvent(new Event('change'));
      // restore display
      const newDisplay = document.createElement('span');
      newDisplay.id = 'dashboard-machine-display';
      newDisplay.textContent = machineNameInput.value.trim() || 'This Machine';
      input.replaceWith(newDisplay);
      updateStatusBar();
    }

    function cancel() {
      const newDisplay = document.createElement('span');
      newDisplay.id = 'dashboard-machine-display';
      newDisplay.textContent = current || 'This Machine';
      input.replaceWith(newDisplay);
    }

    input.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); }
      if (e.key === 'Escape') { e.preventDefault(); cancel(); }
    });
    input.addEventListener('blur', commit);
  }

  heading.addEventListener('click', startEditing);
  heading.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); startEditing(); }
  });

  // Keep display in sync if machine-name input is changed elsewhere (e.g., Account tab)
  machineNameInput.addEventListener('change', refreshDisplay);
  machineNameInput.addEventListener('input', refreshDisplay);
}

initMachineNameRename();

let isSignedIn = false;

// Backup status (keyed by IP/hostname string)
let backupStatusByIp = {};

function appendDebugLogLine(line) {
  if (!debugLogsConsole) return;
  const isAtBottom = (debugLogsConsole.scrollTop + debugLogsConsole.clientHeight) >= (debugLogsConsole.scrollHeight - 10);

  // If the console still has the placeholder, clear it on first real line
  if (debugLogsConsole.childNodes.length === 1) {
    const only = debugLogsConsole.childNodes[0];
    if (only && only.textContent && only.textContent.includes('Waiting for logs')) {
      debugLogsConsole.innerHTML = '';
    }
  }

  const div = document.createElement('div');
  div.textContent = line;
  debugLogsConsole.appendChild(div);

  // Keep DOM from growing unbounded
  const maxLines = 1000;
  while (debugLogsConsole.childNodes.length > maxLines) {
    debugLogsConsole.removeChild(debugLogsConsole.firstChild);
  }

  if (isAtBottom) {
    debugLogsConsole.scrollTop = debugLogsConsole.scrollHeight;
  }
}

async function initDebugLogs() {
  try {
    if (!window.electronAPI || !window.electronAPI.getLogBuffer) return;

    const res = await window.electronAPI.getLogBuffer();
    const lines = Array.isArray(res?.lines) ? res.lines : [];
    if (debugLogsConsole) {
      debugLogsConsole.innerHTML = '';
      if (lines.length === 0) {
        debugLogsConsole.innerHTML = '<div style="color: #888;">No logs yet.</div>';
      } else {
        lines.slice(-300).forEach(appendDebugLogLine);
      }
    }

    if (window.electronAPI.onLogLine) {
      window.electronAPI.onLogLine((line) => appendDebugLogLine(line));
    }

    if (debugLogsClearBtn && window.electronAPI.clearLogBuffer) {
      debugLogsClearBtn.addEventListener('click', async () => {
        await window.electronAPI.clearLogBuffer();
        if (debugLogsConsole) {
          debugLogsConsole.innerHTML = '<div style="color: #888;">Cleared.</div>';
        }
      });
    }

    if (debugLogsSaveBtn && window.electronAPI.exportLogBuffer) {
      debugLogsSaveBtn.addEventListener('click', async () => {
        const result = await window.electronAPI.exportLogBuffer();
        if (result && result.success && result.filePath) {
          showStatus('Saved debug log to: ' + result.filePath, 'info');
        } else if (result && result.canceled) {
          showStatus('Save canceled', 'info');
        } else {
          showStatus('Failed to save debug log', 'error');
        }
      });
    }

    async function refreshCrashInfo() {
      if (!window.electronAPI?.getCrashInfo) return;
      try {
        const info = await window.electronAPI.getCrashInfo();
        if (crashReportsDirEl) crashReportsDirEl.textContent = info.crashReportsDir || '—';
        if (crashDumpsDirEl) crashDumpsDirEl.textContent = info.crashDumpsPath || '—';
        if (lastCrashTimeEl) lastCrashTimeEl.textContent = info.lastCrashTime || 'No crash recorded';
      } catch (e) {
        if (crashReportsDirEl) crashReportsDirEl.textContent = '—';
        if (crashDumpsDirEl) crashDumpsDirEl.textContent = '—';
        if (lastCrashTimeEl) lastCrashTimeEl.textContent = 'Error loading';
      }
    }
    await refreshCrashInfo();
    if (openCrashReportsFolderBtn && window.electronAPI.openCrashReportsFolder) {
      openCrashReportsFolderBtn.addEventListener('click', async () => {
        await window.electronAPI.openCrashReportsFolder();
      });
    }
  } catch (error) {
    console.error('Failed to initialize debug logs:', error);
  }
}

function normalizeControllerIps(ips) {
  if (!Array.isArray(ips)) return [];
  const out = [];
  const seen = new Set();
  ips.forEach((raw) => {
    const v = String(raw || '').trim();
    if (!v) return;
    if (seen.has(v)) return;
    seen.add(v);
    out.push(v);
  });
  return out;
}

function getControllerIpInputs() {
  if (!controllerIpList) return [];
  return Array.from(controllerIpList.querySelectorAll('input[data-controller-ip="true"]'));
}

function getControllerIpsFromUi() {
  return normalizeControllerIps(getControllerIpInputs().map((el) => String(el.value || '').trim()));
}

async function saveControllerAllowlistPreferences() {
  try {
    await window.electronAPI.savePreferences({
      controllerIps: getControllerIpsFromUi()
    });
    showStatus('Controller allowlist saved', 'info');
  } catch (error) {
    console.error('Failed to save controller allowlist:', error);
    showStatus('Failed to save controller allowlist', 'error');
  }
}

function addControllerIpRow(initialValue = '') {
  if (!controllerIpList) return;

  const row = document.createElement('div');
  row.setAttribute('data-controller-row', 'true');
  row.style.display = 'flex';
  row.style.gap = '10px';
  row.style.alignItems = 'center';

  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'input-field';
  input.placeholder = '192.168.1.50';
  input.value = initialValue || '';
  input.setAttribute('data-controller-ip', 'true');
  input.style.flex = '1';

  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.className = 'btn btn-secondary';
  removeBtn.textContent = 'Remove';
  removeBtn.style.padding = '8px 10px';
  removeBtn.style.minWidth = '88px';

  removeBtn.addEventListener('click', async () => {
    const rows = controllerIpList.querySelectorAll('[data-controller-row="true"]');
    if (rows.length <= 1) {
      input.value = '';
      await saveControllerAllowlistPreferences();
      return;
    }
    row.remove();
    await saveControllerAllowlistPreferences();
  });

  input.addEventListener('change', () => {
    saveControllerAllowlistPreferences();
  });

  row.appendChild(input);
  row.appendChild(removeBtn);
  controllerIpList.appendChild(row);
}

function renderControllerIpList(ips = []) {
  if (!controllerIpList) return;
  controllerIpList.innerHTML = '';
  const normalized = Array.isArray(ips) ? ips.map(v => String(v || '')) : [];
  if (normalized.length === 0) {
    addControllerIpRow('');
    return;
  }
  normalized.forEach((ip) => addControllerIpRow(ip));
}

function updateTunnelPinStatusFromPrefs(preferences) {
  if (wanPinScopeSelect && preferences) {
    const s = String(preferences.webUiPinScope || 'tunnel').toLowerCase();
    wanPinScopeSelect.value = s === 'lan' || s === 'both' ? s : 'tunnel';
  }
  if (!wanTunnelPinStatusEl) return;
  if (preferences && preferences.webUiTunnelPinEnabled) {
    const scope = String(preferences.webUiPinScope || 'tunnel').toLowerCase();
    const scopeLabel =
      scope === 'lan' ? 'LAN clients only' : scope === 'both' ? 'tunnel and LAN clients' : 'Cloudflare tunnel clients only';
    wanTunnelPinStatusEl.textContent = `A PIN is set (${scopeLabel}). Matching browsers must unlock the web remote first.`;
  } else {
    wanTunnelPinStatusEl.textContent = 'No PIN is set.';
  }
}

async function saveWanTunnelPinFromDesktop() {
  if (!wanTunnelPinNewInput || !wanTunnelPinConfirmInput) return;
  const a = String(wanTunnelPinNewInput.value || '').trim();
  const b = String(wanTunnelPinConfirmInput.value || '').trim();
  if (a !== b) {
    showStatus('Tunnel PIN and confirmation do not match', 'error');
    return;
  }
  if (!/^[0-9]{4,12}$/.test(a)) {
    showStatus('Tunnel PIN must be 4 to 12 digits', 'error');
    return;
  }
  try {
    await window.electronAPI.savePreferences({ webUiTunnelPin: a });
    wanTunnelPinNewInput.value = '';
    wanTunnelPinConfirmInput.value = '';
    const preferences = await window.electronAPI.getPreferences();
    updateTunnelPinStatusFromPrefs(preferences);
    showStatus('Tunnel PIN saved', 'info');
  } catch (e) {
    console.error('save tunnel pin:', e);
    showStatus('Failed to save tunnel PIN', 'error');
  }
}

async function removeWanTunnelPinFromDesktop() {
  try {
    await window.electronAPI.savePreferences({ webUiTunnelPinClear: true });
    if (wanTunnelPinNewInput) wanTunnelPinNewInput.value = '';
    if (wanTunnelPinConfirmInput) wanTunnelPinConfirmInput.value = '';
    const preferences = await window.electronAPI.getPreferences();
    updateTunnelPinStatusFromPrefs(preferences);
    showStatus('Tunnel PIN removed', 'info');
  } catch (e) {
    console.error('remove tunnel pin:', e);
    showStatus('Failed to remove tunnel PIN', 'error');
  }
}

// Initialize displays
async function initDisplays() {
  try {
    const displays = await window.electronAPI.getDisplays();
    const preferences = await window.electronAPI.getPreferences();
    
    // Clear existing options
    presentationDisplay.innerHTML = '';
    notesDisplay.innerHTML = '';
    
    // Add display options
    displays.forEach(display => {
      const option1 = document.createElement('option');
      option1.value = display.id;
      option1.textContent = display.label + (display.primary ? ' (Primary)' : '');
      presentationDisplay.appendChild(option1);
      
      const option2 = document.createElement('option');
      option2.value = display.id;
      option2.textContent = display.label + (display.primary ? ' (Primary)' : '');
      notesDisplay.appendChild(option2);
    });
    
    // Restore saved preferences or use defaults (coerce to string so <select> matches option values)
    const presId = preferences.presentationDisplayId;
    if (presId !== undefined && presId !== null && presId !== '') {
      presentationDisplay.value = String(presId);
    }

    const notesId = preferences.notesDisplayId;
    if (notesId !== undefined && notesId !== null && notesId !== '') {
      notesDisplay.value = String(notesId);
    } else if (displays.length > 1) {
      // Select different displays by default if available and no preference saved
      notesDisplay.selectedIndex = 1;
    }
    
    // Restore notes layout preference
    const notesLayoutSelect = document.getElementById('notes-layout');
    if (notesLayoutSelect && preferences.notesLayout) {
      notesLayoutSelect.value = preferences.notesLayout;
    }

    const defaultNotesZoomInput = document.getElementById('default-notes-zoom-steps');
    if (defaultNotesZoomInput) {
      const dz = preferences.defaultNotesZoomSteps;
      defaultNotesZoomInput.value = dz !== undefined && dz !== null ? String(dz) : '0';
    }

    const presentationGpuModeSelect = document.getElementById('presentation-gpu-mode');
    if (presentationGpuModeSelect) {
      const allowedGpu = ['default', 'angle-metal', 'angle-gl', 'swiftshader', 'disable-gpu'];
      const m = String(preferences.presentationGpuMode || 'default');
      presentationGpuModeSelect.value = allowedGpu.includes(m) ? m : 'default';
    }
    const presentationNativeFullscreenGroup = document.getElementById('presentation-native-fullscreen-group');
    const presentationNativeFullscreenCheckbox = document.getElementById('presentation-native-fullscreen');
    const isMacUi = /Mac/i.test(navigator.platform || '') || /Mac OS X/.test(navigator.userAgent || '');
    if (presentationNativeFullscreenGroup) {
      presentationNativeFullscreenGroup.style.display = isMacUi ? '' : 'none';
    }
    if (presentationNativeFullscreenCheckbox) {
      presentationNativeFullscreenCheckbox.checked = preferences.presentationNativeFullscreen === true;
    }

    // Restore machine name
    if (preferences.machineName) {
      machineNameInput.value = preferences.machineName;
    }

    // Restore port preferences
    if (preferences.apiPort) {
      apiPortInput.value = preferences.apiPort;
    } else {
      apiPortInput.value = '9595'; // Default
    }

    if (preferences.webUiPort) {
      webUiPortInput.value = preferences.webUiPort;
    } else {
      webUiPortInput.value = '80'; // Default
    }

    if (webUiUseHttpsCheckbox) {
      webUiUseHttpsCheckbox.checked = preferences.webUiUseHttps === true;
      updateHttpsCertKeyVisibility();
    }
    if (webUiCertPathInput) webUiCertPathInput.value = preferences.webUiCertPath || '';
    if (webUiKeyPathInput) webUiKeyPathInput.value = preferences.webUiKeyPath || '';
    
    // Restore logging preferences
    if (verboseLoggingCheckbox) {
      verboseLoggingCheckbox.checked = preferences.verboseLogging === true;
    }
    if (webUiDebugConsoleEnabledCheckbox) {
      webUiDebugConsoleEnabledCheckbox.checked = preferences.webUiDebugConsoleEnabled === true;
    }
    if (allowArbitraryUrlCheckbox) {
      allowArbitraryUrlCheckbox.checked = preferences.allowArbitraryUrl === true;
    }
    const savedBgColor = /^#[0-9a-fA-F]{6}$/.test(preferences.genericUrlBackgroundColor)
      ? preferences.genericUrlBackgroundColor : '#000000';
    if (genericUrlBgColorPicker) genericUrlBgColorPicker.value = savedBgColor;
    if (genericUrlBgColorHex) genericUrlBgColorHex.value = savedBgColor;

    // Restore keyboard shortcut preferences
    if (keyboardShortcutPresetSelect) {
      const validPresets = ['cmd+arrow', 'alt+arrow', 'cmd+shift+arrow'];
      const preset = preferences.keyboardShortcutPreset;
      keyboardShortcutPresetSelect.value = validPresets.includes(preset) ? preset : 'cmd+arrow';
    }
    if (keyboardShortcutsDefaultEnabledCheckbox) {
      keyboardShortcutsDefaultEnabledCheckbox.checked = preferences.keyboardShortcutsDefaultEnabled === true;
    }

    // Restore Web UI appearance (theme + logo + custom CSS path)
    if (webUiThemeSelect) {
      const theme = preferences.webUiTheme || 'original';
      webUiThemeSelect.value = ['original', 'light', 'dark', 'max', 'touch', 'thumb'].includes(theme) ? theme : 'original';
    }
    if (webUiLogoPathInput) {
      webUiLogoPathInput.value = preferences.webUiLogoPath || '';
    }
    if (webUiCustomCssPathInput) {
      webUiCustomCssPathInput.value = preferences.webUiCustomCssPath || '';
    }

    updateTunnelPinStatusFromPrefs(preferences);
    
    // Restore primary/backup mode
    const mode = preferences.primaryBackupMode || 'standalone';
    if (mode === 'primary') {
      modePrimary.checked = true;
      backupConfig.style.display = 'block';
    } else if (mode === 'backup') {
      modeBackup.checked = true;
      backupConfig.style.display = 'none';
    } else {
      modeStandalone.checked = true;
      backupConfig.style.display = 'none';
    }
    
    // Restore backup configuration
    if (preferences.backupPort) {
      backupPortInput.value = preferences.backupPort;
    } else {
      backupPortInput.value = '9595'; // Default
    }

    // Restore backup IP list (unlimited). Fallback to legacy fields if present.
    const legacyIps = [preferences.backupIp1, preferences.backupIp2, preferences.backupIp3].filter(Boolean);
    const backupIps = Array.isArray(preferences.backupIps) ? preferences.backupIps : legacyIps;
    renderBackupIpList(backupIps);
    refreshBackupStatusBadges();

    // Restore controller allowlist (desktop-only)
    const controllerIps = Array.isArray(preferences.controllerIps) ? preferences.controllerIps : [];
    renderControllerIpList(controllerIps);

    // Restore PerfectCue settings
    if (perfectCueEnabledCheckbox) {
      perfectCueEnabledCheckbox.checked = preferences.perfectCueEnabled === true;
    }
    // preferences.perfectCuePorts is already PortConfig[] after loadPreferences() normalization
    const perfectCuePorts = Array.isArray(preferences.perfectCuePorts) ? preferences.perfectCuePorts : [];
    renderPerfectCuePortList(perfectCuePorts);
    
    // Save preferences when selections change
    presentationDisplay.addEventListener('change', saveMonitorPreferences);
    notesDisplay.addEventListener('change', saveMonitorPreferences);
    const notesLayoutEl = document.getElementById('notes-layout');
    if (notesLayoutEl) notesLayoutEl.addEventListener('change', saveMonitorPreferences);
    const relaunchNotesBtn = document.getElementById('btn-relaunch-speaker-notes');
    if (relaunchNotesBtn) {
      relaunchNotesBtn.addEventListener('click', async () => {
        const originalText = relaunchNotesBtn.textContent;
        relaunchNotesBtn.disabled = true;
        relaunchNotesBtn.textContent = 'Relaunching...';
        try {
          const result = await window.electronAPI.relaunchSpeakerNotes();
          if (result && result.success) {
            showStatus(result.message || 'Speaker notes relaunched', 'info');
          } else {
            showStatus(result?.error || 'Relaunch failed', 'error');
          }
        } catch (err) {
          showStatus(err.message || 'Relaunch failed', 'error');
        } finally {
          relaunchNotesBtn.disabled = false;
          relaunchNotesBtn.textContent = originalText;
        }
      });
    }
    const defaultNotesZoomEl = document.getElementById('default-notes-zoom-steps');
    if (defaultNotesZoomEl) defaultNotesZoomEl.addEventListener('change', saveMonitorPreferences);
    const savePresentationRenderingBtn = document.getElementById('save-presentation-rendering-btn');
    if (savePresentationRenderingBtn) {
      savePresentationRenderingBtn.addEventListener('click', savePresentationRenderingPreferences);
    }
    machineNameInput.addEventListener('change', saveMachineName);
    apiPortInput.addEventListener('change', savePortPreferences);
    webUiPortInput.addEventListener('change', savePortPreferences);
    if (webUiUseHttpsCheckbox) {
      webUiUseHttpsCheckbox.addEventListener('change', () => {
        if (webUiUseHttpsCheckbox.checked && webUiPortInput) {
          const port = parseInt(webUiPortInput.value, 10);
          if (port === 80 || !port) {
            webUiPortInput.value = '443';
          }
        }
        updateHttpsCertKeyVisibility();
        saveHttpsPreferences();
      });
    }
    if (webUiCertChooseBtn && window.electronAPI.showOpenCertDialog) {
      webUiCertChooseBtn.addEventListener('click', async () => {
        try {
          const result = await window.electronAPI.showOpenCertDialog();
          if (result && result.filePath && webUiCertPathInput) {
            webUiCertPathInput.value = result.filePath;
            saveHttpsPreferences();
          }
        } catch (e) {
          console.error('Open cert dialog failed:', e);
          showStatus('Could not open file dialog', 'error');
        }
      });
    }
    if (webUiCertClearBtn && webUiCertPathInput) {
      webUiCertClearBtn.addEventListener('click', () => {
        webUiCertPathInput.value = '';
        saveHttpsPreferences();
      });
    }
    if (webUiKeyChooseBtn && window.electronAPI.showOpenKeyDialog) {
      webUiKeyChooseBtn.addEventListener('click', async () => {
        try {
          const result = await window.electronAPI.showOpenKeyDialog();
          if (result && result.filePath && webUiKeyPathInput) {
            webUiKeyPathInput.value = result.filePath;
            saveHttpsPreferences();
          }
        } catch (e) {
          console.error('Open key dialog failed:', e);
          showStatus('Could not open file dialog', 'error');
        }
      });
    }
    if (webUiKeyClearBtn && webUiKeyPathInput) {
      webUiKeyClearBtn.addEventListener('click', () => {
        webUiKeyPathInput.value = '';
        saveHttpsPreferences();
      });
    }
    if (verboseLoggingCheckbox) {
      verboseLoggingCheckbox.addEventListener('change', saveLoggingPreferences);
    }
    if (webUiDebugConsoleEnabledCheckbox) {
      webUiDebugConsoleEnabledCheckbox.addEventListener('change', saveWebUiDebugConsolePreference);
    }
    if (allowArbitraryUrlCheckbox) {
      allowArbitraryUrlCheckbox.addEventListener('change', saveAllowArbitraryUrlPreference);
    }
    if (genericUrlBgColorPicker) {
      genericUrlBgColorPicker.addEventListener('input', () => {
        if (genericUrlBgColorHex) genericUrlBgColorHex.value = genericUrlBgColorPicker.value;
      });
      genericUrlBgColorPicker.addEventListener('change', saveGenericUrlBackgroundColor);
    }
    if (genericUrlBgColorHex) {
      genericUrlBgColorHex.addEventListener('change', () => {
        const v = genericUrlBgColorHex.value.trim();
        if (/^#[0-9a-fA-F]{6}$/.test(v)) {
          if (genericUrlBgColorPicker) genericUrlBgColorPicker.value = v;
          saveGenericUrlBackgroundColor();
        }
      });
    }

    // Web UI appearance: logo and CSS file pickers, save
    if (webUiLogoChooseBtn && window.electronAPI.showOpenLogoDialog) {
      webUiLogoChooseBtn.addEventListener('click', async () => {
        try {
          const result = await window.electronAPI.showOpenLogoDialog();
          if (result && result.filePath && webUiLogoPathInput) {
            webUiLogoPathInput.value = result.filePath;
          }
        } catch (e) {
          console.error('Open logo dialog failed:', e);
          showStatus('Could not open file dialog', 'error');
        }
      });
    }
    if (webUiLogoClearBtn && webUiLogoPathInput) {
      webUiLogoClearBtn.addEventListener('click', () => {
        webUiLogoPathInput.value = '';
      });
    }
    if (webUiCustomCssChooseBtn && window.electronAPI.showOpenCssDialog) {
      webUiCustomCssChooseBtn.addEventListener('click', async () => {
        try {
          const result = await window.electronAPI.showOpenCssDialog();
          if (result && result.filePath && webUiCustomCssPathInput) {
            webUiCustomCssPathInput.value = result.filePath;
          }
        } catch (e) {
          console.error('Open CSS dialog failed:', e);
          showStatus('Could not open file dialog', 'error');
        }
      });
    }
    if (webUiCustomCssClearBtn && webUiCustomCssPathInput) {
      webUiCustomCssClearBtn.addEventListener('click', () => {
        webUiCustomCssPathInput.value = '';
      });
    }
    if (saveWebUiAppearanceBtn) {
      saveWebUiAppearanceBtn.addEventListener('click', saveWebUiAppearance);
    }
    if (saveKeyboardShortcutsBtn) {
      saveKeyboardShortcutsBtn.addEventListener('click', saveKeyboardShortcutSettings);
    }
    if (webUiDownloadCssTemplateBtn && window.electronAPI.downloadCssTemplate) {
      webUiDownloadCssTemplateBtn.addEventListener('click', async () => {
        try {
          const result = await window.electronAPI.downloadCssTemplate();
          if (result && result.success && result.filePath) {
            showStatus('CSS template saved to: ' + result.filePath, 'info');
          } else if (result && result.canceled) {
            showStatus('Save canceled', 'info');
          } else {
            showStatus(result && result.error ? result.error : 'Failed to save template', 'error');
          }
        } catch (e) {
          console.error('Download CSS template failed:', e);
          showStatus('Failed to save CSS template', 'error');
        }
      });
    }
    
    // Primary/Backup mode change handlers
    modePrimary.addEventListener('change', () => {
      if (modePrimary.checked) {
        backupConfig.style.display = 'block';
        savePrimaryBackupPreferences();
        updateStatusBar();
      }
    });

    modeBackup.addEventListener('change', () => {
      if (modeBackup.checked) {
        backupConfig.style.display = 'none';
        savePrimaryBackupPreferences();
        updateStatusBar();
      }
    });

    modeStandalone.addEventListener('change', () => {
      if (modeStandalone.checked) {
        backupConfig.style.display = 'none';
        savePrimaryBackupPreferences();
        updateStatusBar();
      }
    });
    
    // Backup configuration change handlers
    backupPortInput.addEventListener('change', savePrimaryBackupPreferences);
    if (addBackupIpBtn) {
      addBackupIpBtn.addEventListener('click', async () => {
        addBackupIpRow('');
        const inputs = getBackupIpInputs();
        if (inputs.length) inputs[inputs.length - 1].focus();
        await savePrimaryBackupPreferences();
      });
    }

    if (addControllerIpBtn) {
      addControllerIpBtn.addEventListener('click', async () => {
        addControllerIpRow('');
        const inputs = getControllerIpInputs();
        if (inputs.length) inputs[inputs.length - 1].focus();
        await saveControllerAllowlistPreferences();
      });
    }

    if (addPerfectCuePortBtn) {
      addPerfectCuePortBtn.addEventListener('click', () => {
        addPerfectCuePortRow({ port: '', name: '', enabled: true });
        const inputs = perfectCuePortList ? perfectCuePortList.querySelectorAll('input[data-perfectcue-port="true"]') : [];
        if (inputs.length) inputs[inputs.length - 1].focus();
      });
    }
    if (savePerfectCueBtn) {
      savePerfectCueBtn.addEventListener('click', savePerfectCuePrefs);
    }
    
    // Load and display network info
    await updateNetworkInfo();

    // WAN Access (Cloudflare Quick Tunnel)
    await loadTunnelStatus();
    if (wanEnabledCheckbox && window.electronAPI.setTunnelEnabled) {
      wanEnabledCheckbox.addEventListener('change', async () => {
        const enabled = wanEnabledCheckbox.checked;
        if (wanStatusRow) wanStatusRow.style.display = enabled ? '' : 'none';
        if (wanUrlDisplay) wanUrlDisplay.value = enabled ? 'Connecting\u2026' : '';
        try {
          await window.electronAPI.setTunnelEnabled(enabled);
        } catch (e) {
          console.error('setTunnelEnabled failed:', e);
          showStatus('Could not update WAN tunnel setting', 'error');
        }
        updateStatusBar();
        syncDashboardWan();
      });
    }
    if (wanCopyBtn && wanUrlDisplay) {
      wanCopyBtn.addEventListener('click', () => {
        const url = wanUrlDisplay.value;
        if (!url || url === 'Connecting…') {
          showStatus('Tunnel URL not ready yet', 'error');
          return;
        }
        navigator.clipboard.writeText(url)
          .then(() => showStatus('Tunnel URL copied', 'info'))
          .catch(() => showStatus('Copy failed', 'error'));
      });
    }
    if (window.electronAPI.onTunnelUrlChanged) {
      window.electronAPI.onTunnelUrlChanged((url) => {
        if (wanUrlDisplay) wanUrlDisplay.value = url || '';
        if (dashboardWanUrlDisplay) dashboardWanUrlDisplay.value = wanUrlDisplay ? (wanUrlDisplay.value) : (url || '');
        if (url) showStatus('WAN tunnel connected', 'info');
      });
    }

    if (wanTunnelPinSaveBtn) {
      wanTunnelPinSaveBtn.addEventListener('click', () => {
        saveWanTunnelPinFromDesktop();
      });
    }
    if (wanTunnelPinRemoveBtn) {
      wanTunnelPinRemoveBtn.addEventListener('click', () => {
        removeWanTunnelPinFromDesktop();
      });
    }
    if (wanPinScopeSelect) {
      wanPinScopeSelect.addEventListener('change', async () => {
        const v = wanPinScopeSelect.value;
        const scope = v === 'lan' || v === 'both' ? v : 'tunnel';
        try {
          await window.electronAPI.savePreferences({ webUiPinScope: scope });
          const preferences = await window.electronAPI.getPreferences();
          updateTunnelPinStatusFromPrefs(preferences);
          showStatus('PIN scope saved', 'info');
        } catch (e) {
          console.error('save pin scope:', e);
          showStatus('Failed to save PIN scope', 'error');
        }
      });
    }

    // Start polling backup status if in primary mode
    if (mode === 'primary') {
      startBackupStatusPolling();
    }
    
    // Load preset presentations
    await loadPresets();
    
    // Load stagetimer settings
    await loadStagetimerSettings();
    
    // Set up event handlers for presets
    if (addPresetBtn) {
      addPresetBtn.addEventListener('click', () => {
        addPresetRow('');
        const inputs = presetList ? presetList.querySelectorAll('input[data-preset-url="true"]') : [];
        if (inputs.length) inputs[inputs.length - 1].focus();
      });
    }
    savePresetsBtn.addEventListener('click', savePresets);
    loadPresetsBtn.addEventListener('click', loadPresets);
    
    // Set up event handlers for stagetimer
    saveStagetimerBtn.addEventListener('click', saveStagetimerSettings);
    loadStagetimerBtn.addEventListener('click', loadStagetimerSettings);

    // Speaker notes capture (clean text via IPC - no fetch/port needed)
    async function refreshSpeakerNotesCapture() {
      if (!speakerNotesCapture) return;
      const warnEl = document.getElementById('notes-encoding-warning-desktop');
      try {
        const data = await window.electronAPI.getSpeakerNotes();
        if (data.success && data.notes) {
          speakerNotesCapture.textContent = data.notes;
          if (warnEl) warnEl.style.display = data.encodingIssuesDetected ? 'block' : 'none';
        } else {
          speakerNotesCapture.textContent = data.error || 'No speaker notes window open.';
          if (warnEl) warnEl.style.display = 'none';
        }
      } catch (err) {
        speakerNotesCapture.textContent = 'Could not load notes. Is speaker notes open?';
        if (warnEl) warnEl.style.display = 'none';
      }
    }
    if (speakerNotesRefreshBtn) speakerNotesRefreshBtn.addEventListener('click', refreshSpeakerNotesCapture);
    if (speakerNotesCopyBtn && speakerNotesCapture) {
      speakerNotesCopyBtn.addEventListener('click', () => {
        const text = speakerNotesCapture.textContent || '';
        if (!text || text.startsWith('No speaker notes') || text.startsWith('Could not load')) {
          showStatus('Nothing to copy. Open speaker notes and click Refresh.', 'error');
          return;
        }
        navigator.clipboard.writeText(text).then(() => showStatus('Speaker notes copied to clipboard', 'info')).catch(() => showStatus('Copy failed', 'error'));
      });
    }
    refreshSpeakerNotesCapture();

    updateStatusBar();
  } catch (error) {
    showStatus('Failed to load displays', 'error');
  }
}

function normalizeBackupIps(ips) {
  if (!Array.isArray(ips)) return [];
  const out = [];
  const seen = new Set();
  for (const raw of ips) {
    const v = String(raw || '').trim();
    if (!v) continue;
    if (seen.has(v)) continue;
    seen.add(v);
    out.push(v);
  }
  return out;
}

function getBackupIpInputs() {
  if (!backupIpList) return [];
  return Array.from(backupIpList.querySelectorAll('input[data-backup-ip="true"]'));
}

function getBackupIpsFromUi({ includeEmpty = false } = {}) {
  const values = getBackupIpInputs().map((el) => String(el.value || '').trim());
  if (includeEmpty) return values;
  return normalizeBackupIps(values);
}

function setBackupStatusBadge(el, ip) {
  if (!el) return;
  const v = String(ip || '').trim();
  const status = v ? backupStatusByIp[v] : null;

  if (!v) {
    el.textContent = '-';
    el.style.background = 'transparent';
    el.style.color = 'var(--text-secondary)';
    return;
  }
  if (status === 'connected') {
    el.textContent = 'Connected';
    el.style.background = '#4caf50';
    el.style.color = 'white';
    return;
  }
  if (status === 'disconnected') {
    el.textContent = 'Disconnected';
    el.style.background = '#f44336';
    el.style.color = 'white';
    return;
  }
  el.textContent = 'Checking...';
  el.style.background = '#ff9800';
  el.style.color = 'white';
}

function refreshBackupStatusBadges() {
  if (!backupIpList) return;
  const rows = Array.from(backupIpList.querySelectorAll('[data-backup-row="true"]'));
  rows.forEach((row) => {
    const input = row.querySelector('input[data-backup-ip="true"]');
    const badge = row.querySelector('span[data-backup-status="true"]');
    setBackupStatusBadge(badge, input ? input.value : '');
  });
}

function addBackupIpRow(initialValue = '') {
  if (!backupIpList) return;

  const row = document.createElement('div');
  row.setAttribute('data-backup-row', 'true');
  row.style.display = 'flex';
  row.style.gap = '10px';
  row.style.alignItems = 'center';
  row.style.width = '100%';
  row.style.minWidth = '0';

  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'input-field';
  input.placeholder = '192.168.1.100 or hostname';
  input.value = initialValue || '';
  input.setAttribute('data-backup-ip', 'true');
  input.style.flex = '1 1 0%';
  input.style.width = '100%';
  input.style.minWidth = '140px';

  const badge = document.createElement('span');
  badge.setAttribute('data-backup-status', 'true');
  badge.style.fontSize = '12px';
  badge.style.padding = '4px 8px';
  badge.style.borderRadius = '4px';
  badge.style.minWidth = '90px';
  badge.style.textAlign = 'center';

  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.className = 'btn btn-secondary';
  removeBtn.textContent = 'Remove';
  removeBtn.style.padding = '8px 10px';
  removeBtn.style.minWidth = '88px';

  removeBtn.addEventListener('click', async () => {
    // If it's the last row, just clear it (cleaner UX)
    const rows = backupIpList.querySelectorAll('[data-backup-row="true"]');
    if (rows.length <= 1) {
      input.value = '';
      await savePrimaryBackupPreferences();
      refreshBackupStatusBadges();
      return;
    }
    row.remove();
    await savePrimaryBackupPreferences();
    refreshBackupStatusBadges();
  });

  input.addEventListener('change', () => {
    savePrimaryBackupPreferences();
    refreshBackupStatusBadges();
  });

  row.appendChild(input);
  row.appendChild(badge);
  row.appendChild(removeBtn);
  backupIpList.appendChild(row);

  setBackupStatusBadge(badge, input.value);
}

function renderBackupIpList(ips = []) {
  if (!backupIpList) return;
  backupIpList.innerHTML = '';

  const normalized = Array.isArray(ips) ? ips.map(v => String(v || '')) : [];
  // Keep UI clean: show 1 row by default in primary mode
  if (normalized.length === 0) {
    addBackupIpRow('');
    return;
  }
  normalized.forEach((ip) => addBackupIpRow(ip));
}

function addPerfectCuePortRow(config = {}) {
  if (!perfectCuePortList) return;
  const { port = '', name = '', enabled = true, adapter = 'dsan' } = config;

  const row = document.createElement('div');
  row.setAttribute('data-perfectcue-row', 'true');
  row.style.cssText = 'display:flex;gap:8px;align-items:center;width:100%;min-width:0;flex-wrap:wrap;';

  // Name input
  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.className = 'input-field';
  nameInput.placeholder = 'Name (optional)';
  nameInput.maxLength = 40;
  nameInput.value = name;
  nameInput.setAttribute('data-perfectcue-name', 'true');
  nameInput.style.cssText = 'flex:2 1 120px;min-width:100px;';

  // Port number input
  const portInput = document.createElement('input');
  portInput.type = 'number';
  portInput.className = 'input-field';
  portInput.placeholder = '8899';
  portInput.min = '1024';
  portInput.max = '65535';
  portInput.value = port || '';
  portInput.setAttribute('data-perfectcue-port', 'true');
  portInput.style.cssText = 'flex:1 1 90px;min-width:80px;';

  const adapterSelect = document.createElement('select');
  adapterSelect.className = 'input-field';
  adapterSelect.setAttribute('data-perfectcue-adapter', 'true');
  adapterSelect.setAttribute('aria-label', 'Converter type');
  adapterSelect.style.cssText = 'flex:1 1 130px;min-width:120px;';
  [['dsan', 'DSAN'], ['waveshare', 'WaveShare']].forEach(([value, label]) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = label;
    adapterSelect.appendChild(opt);
  });
  adapterSelect.value = adapter === 'waveshare' ? 'waveshare' : 'dsan';

  // Enabled checkbox label
  const enabledLabel = document.createElement('label');
  enabledLabel.style.cssText = 'display:flex;align-items:center;gap:4px;cursor:pointer;white-space:nowrap;';

  const enabledCheckbox = document.createElement('input');
  enabledCheckbox.type = 'checkbox';
  enabledCheckbox.checked = enabled !== false;
  enabledCheckbox.setAttribute('data-perfectcue-enabled', 'true');

  const enabledText = document.createElement('span');
  enabledText.textContent = 'Enabled';
  enabledText.style.fontSize = '13px';

  enabledLabel.appendChild(enabledCheckbox);
  enabledLabel.appendChild(enabledText);

  // Disable checkbox if port value is not yet valid (unsaved new row)
  const syncCheckboxToPort = () => {
    const v = parseInt(portInput.value, 10);
    enabledCheckbox.disabled = !(Number.isInteger(v) && v >= 1024 && v <= 65535);
  };
  syncCheckboxToPort();
  portInput.addEventListener('input', syncCheckboxToPort);

  // Immediate IPC on toggle
  enabledCheckbox.addEventListener('change', async () => {
    const portVal = parseInt(portInput.value, 10);
    if (!Number.isInteger(portVal) || portVal < 1024 || portVal > 65535) return;
    const newChecked = enabledCheckbox.checked;
    try {
      const result = await window.electronAPI.togglePerfectCuePort(portVal, newChecked);
      if (result && !result.success) {
        enabledCheckbox.checked = !newChecked; // revert
        showStatus('Save PerfectCue settings first before toggling', 'error');
      }
    } catch (err) {
      enabledCheckbox.checked = !newChecked; // revert
      console.error('[PerfectCue] Failed to toggle port:', err);
    }
  });

  // Remove button
  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.className = 'btn btn-secondary';
  removeBtn.textContent = 'Remove';
  removeBtn.style.cssText = 'padding:8px 10px;min-width:80px;';

  removeBtn.addEventListener('click', () => {
    const rows = perfectCuePortList.querySelectorAll('[data-perfectcue-row="true"]');
    if (rows.length <= 1) {
      portInput.value = '';
      nameInput.value = '';
      adapterSelect.value = 'dsan';
      enabledCheckbox.checked = true;
      return;
    }
    row.remove();
  });

  row.appendChild(nameInput);
  row.appendChild(portInput);
  row.appendChild(adapterSelect);
  row.appendChild(enabledLabel);
  row.appendChild(removeBtn);
  perfectCuePortList.appendChild(row);
}

function renderPerfectCuePortList(portConfigs = []) {
  if (!perfectCuePortList) return;
  perfectCuePortList.innerHTML = '';
  const configs = Array.isArray(portConfigs) ? portConfigs : [];
  if (configs.length === 0) {
    addPerfectCuePortRow({ port: 8899, name: '', enabled: true, adapter: 'dsan' });
    return;
  }
  configs.forEach(c => addPerfectCuePortRow(c));
}

function getPerfectCuePortsFromUi() {
  if (!perfectCuePortList) return [];
  return Array.from(perfectCuePortList.querySelectorAll('[data-perfectcue-row="true"]'))
    .map(row => {
      const portEl = row.querySelector('[data-perfectcue-port="true"]');
      const nameEl = row.querySelector('[data-perfectcue-name="true"]');
      const adapterEl = row.querySelector('[data-perfectcue-adapter="true"]');
      const enabledEl = row.querySelector('[data-perfectcue-enabled="true"]');
      const port = portEl ? parseInt(portEl.value, 10) : NaN;
      const name = nameEl ? nameEl.value.trim() : '';
      const adapter = adapterEl && adapterEl.value === 'waveshare' ? 'waveshare' : 'dsan';
      const enabled = enabledEl ? enabledEl.checked : true;
      return { port, name, enabled, adapter };
    })
    .filter(c => Number.isInteger(c.port) && c.port >= 1024 && c.port <= 65535);
}

async function savePerfectCuePrefs() {
  try {
    const enabled = perfectCueEnabledCheckbox ? perfectCueEnabledCheckbox.checked : false;
    const ports = getPerfectCuePortsFromUi();
    await window.electronAPI.savePreferences({ perfectCueEnabled: enabled, perfectCuePorts: ports });
    showStatus('PerfectCue settings saved', 'info');
  } catch (err) {
    console.error('Failed to save PerfectCue preferences:', err);
    showStatus('Failed to save PerfectCue settings', 'error');
  }
}

async function loadTunnelStatus() {
  if (!wanEnabledCheckbox || !window.electronAPI.getTunnelStatus) return;
  try {
    const status = await window.electronAPI.getTunnelStatus();
    wanEnabledCheckbox.checked = !!status.enabled;
    if (wanStatusRow) {
      wanStatusRow.style.display = status.enabled ? '' : 'none';
    }
    if (wanUrlDisplay) {
      wanUrlDisplay.value = status.url || (status.running ? 'Connecting\u2026' : '');
    }
    updateStatusBar();
    syncDashboardWan();
  } catch (err) {
    console.error('Failed to load tunnel status:', err);
  }
}

function renderUrlListInto(containerId, networkInfo, port, protocol) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = '';
  if (!networkInfo || networkInfo.length === 0) {
    el.innerHTML = '<div class="url-item">No network interfaces found</div>';
    return;
  }
  networkInfo.forEach(ip => {
    const item = document.createElement('div');
    item.className = 'url-item' + (ip.internal ? ' internal' : '');
    const url = `${protocol}://${ip.address}:${port}`;
    item.textContent = url;
    if (ip.internal) {
      item.title = 'Localhost/internal interface';
    } else {
      item.style.cursor = 'pointer';
      item.title = 'Click to open';
      item.addEventListener('click', () => window.electronAPI.openExternal(url));
    }
    el.appendChild(item);
  });
}

// Update network information display
async function updateNetworkInfo() {
  try {
    const networkInfo = await window.electronAPI.getNetworkInfo();
    const preferences = await window.electronAPI.getPreferences();

    const apiPort = preferences.apiPort || 9595;
    let webUiPort = preferences.webUiPort || 80;
    if (preferences.webUiUseHttps && webUiPort === 80) webUiPort = 443;
    const webUiProtocol = preferences.webUiUseHttps ? 'https' : 'http';

    renderUrlListInto('api-urls',              networkInfo, apiPort,   'http');
    renderUrlListInto('web-ui-urls',           networkInfo, webUiPort, webUiProtocol);
    renderUrlListInto('dashboard-api-urls',    networkInfo, apiPort,   'http');
    renderUrlListInto('dashboard-web-ui-urls', networkInfo, webUiPort, webUiProtocol);
  } catch (error) {
    console.error('Failed to load network info:', error);
    ['api-urls', 'web-ui-urls', 'dashboard-api-urls', 'dashboard-web-ui-urls'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.innerHTML = '<div class="url-item">Error loading network info</div>';
    });
  }
}

async function savePresentationRenderingPreferences() {
  try {
    const gpuEl = document.getElementById('presentation-gpu-mode');
    const nativeEl = document.getElementById('presentation-native-fullscreen');
    const mode = gpuEl ? String(gpuEl.value || 'default') : 'default';
    const allowedGpu = ['default', 'angle-metal', 'angle-gl', 'swiftshader', 'disable-gpu'];
    const presentationGpuMode = allowedGpu.includes(mode) ? mode : 'default';
    await window.electronAPI.savePreferences({
      presentationGpuMode,
      presentationNativeFullscreen: nativeEl ? nativeEl.checked === true : false
    });
    showStatus('Presentation rendering saved. Restart the app for GPU mode to take effect.', 'info');
  } catch (error) {
    console.error('Failed to save presentation rendering:', error);
    showStatus('Failed to save presentation rendering', 'error');
  }
}

// Save monitor preferences
async function saveMonitorPreferences() {
  try {
    const notesLayoutSelect = document.getElementById('notes-layout');
    const defaultNotesZoomEl = document.getElementById('default-notes-zoom-steps');
    let defaultNotesZoomSteps = 0;
    if (defaultNotesZoomEl) {
      const parsed = parseInt(defaultNotesZoomEl.value, 10);
      defaultNotesZoomSteps = Number.isNaN(parsed) ? 0 : parsed;
    }
    await window.electronAPI.savePreferences({
      presentationDisplayId: parseInt(presentationDisplay.value, 10),
      notesDisplayId: parseInt(notesDisplay.value, 10),
      notesLayout: notesLayoutSelect ? notesLayoutSelect.value : 'hide',
      defaultNotesZoomSteps
    });
  } catch (error) {
    console.error('Failed to save preferences:', error);
  }
}

// Save machine name
async function saveMachineName() {
  try {
    const machineName = machineNameInput.value.trim();
    await window.electronAPI.savePreferences({
      machineName: machineName || null
    });
    showStatus('Machine name saved', 'info');
    updateStatusBar();
  } catch (error) {
    console.error('Failed to save machine name:', error);
    showStatus('Failed to save machine name', 'error');
  }
}

// Save port preferences
function updateHttpsCertKeyVisibility() {
  if (!webUiHttpsCertGroup || !webUiHttpsKeyGroup || !webUiUseHttpsCheckbox) return;
  const show = webUiUseHttpsCheckbox.checked;
  webUiHttpsCertGroup.style.display = show ? 'block' : 'none';
  webUiHttpsKeyGroup.style.display = show ? 'block' : 'none';
}

async function saveHttpsPreferences() {
  try {
    const useHttps = webUiUseHttpsCheckbox ? webUiUseHttpsCheckbox.checked === true : false;
    const port = webUiPortInput ? (parseInt(webUiPortInput.value, 10) || (useHttps ? 443 : 80)) : undefined;
    const prefs = {
      webUiUseHttps: useHttps,
      webUiCertPath: webUiCertPathInput ? (webUiCertPathInput.value || '').trim() || null : null,
      webUiKeyPath: webUiKeyPathInput ? (webUiKeyPathInput.value || '').trim() || null : null
    };
    if (port !== undefined) prefs.webUiPort = port;
    await window.electronAPI.savePreferences(prefs);
    showStatus('HTTPS settings saved. Restart the app for Web UI to use HTTPS.', 'info');
  } catch (error) {
    console.error('Failed to save HTTPS preferences:', error);
    showStatus('Failed to save HTTPS settings', 'error');
  }
}

async function savePortPreferences() {
  try {
    const apiPort = parseInt(apiPortInput.value, 10);
    const webUiPort = parseInt(webUiPortInput.value, 10);
    
    // Validate ports
    if (isNaN(apiPort) || apiPort < 1024 || apiPort > 65535) {
      showStatus('API port must be between 1024 and 65535', 'error');
      return;
    }
    
    if (isNaN(webUiPort) || webUiPort < 1 || webUiPort > 65535) {
      showStatus('Web UI port must be between 1 and 65535', 'error');
      return;
    }
    
    const prefsToSave = {
      apiPort: apiPort,
      webUiPort: webUiPort
    };
    // Keep backup target port aligned with API port on primaries (health checks + sendToBackups use both).
    if (modePrimary.checked) {
      prefsToSave.backupPort = apiPort;
    }
    await window.electronAPI.savePreferences(prefsToSave);
    updateStatusBar();
    if (modePrimary.checked) {
      backupPortInput.value = String(apiPort);
    }

    // Update network info display with new ports
    await updateNetworkInfo();

    showStatus('Ports saved. Please restart the app for changes to take effect.', 'info');
  } catch (error) {
    console.error('Failed to save port preferences:', error);
    showStatus('Failed to save port preferences', 'error');
  }
}

// Save logging preferences
async function saveLoggingPreferences() {
  try {
    await window.electronAPI.savePreferences({
      verboseLogging: verboseLoggingCheckbox && verboseLoggingCheckbox.checked === true
    });
    showStatus('Logging settings saved', 'info');
  } catch (error) {
    console.error('Failed to save logging preferences:', error);
    showStatus('Failed to save logging preferences', 'error');
  }
}

async function saveWebUiDebugConsolePreference() {
  try {
    await window.electronAPI.savePreferences({
      webUiDebugConsoleEnabled: webUiDebugConsoleEnabledCheckbox && webUiDebugConsoleEnabledCheckbox.checked === true
    });
    showStatus('Web UI debug console setting saved', 'info');
  } catch (error) {
    console.error('Failed to save Web UI debug console preference:', error);
    showStatus('Failed to save Web UI debug console preference', 'error');
  }
}

async function saveAllowArbitraryUrlPreference() {
  try {
    await window.electronAPI.savePreferences({
      allowArbitraryUrl: allowArbitraryUrlCheckbox && allowArbitraryUrlCheckbox.checked === true
    });
    showStatus('Web remote features saved', 'info');
  } catch (error) {
    console.error('Failed to save allow arbitrary URL preference:', error);
    showStatus('Failed to save web remote features setting', 'error');
  }
}

async function saveGenericUrlBackgroundColor() {
  const color = genericUrlBgColorPicker ? genericUrlBgColorPicker.value : '#000000';
  try {
    await window.electronAPI.savePreferences({ genericUrlBackgroundColor: color });
    showStatus('Background color saved', 'info');
  } catch (error) {
    console.error('Failed to save background color:', error);
    showStatus('Failed to save background color', 'error');
  }
}

async function saveWebUiAppearance() {
  try {
    const theme = webUiThemeSelect ? webUiThemeSelect.value : 'original';
    const logoPath = webUiLogoPathInput ? (webUiLogoPathInput.value || '').trim() : '';
    const path = webUiCustomCssPathInput ? (webUiCustomCssPathInput.value || '').trim() : '';
    await window.electronAPI.savePreferences({
      webUiTheme: ['original', 'light', 'dark', 'max', 'touch', 'thumb'].includes(theme) ? theme : 'original',
      webUiLogoPath: logoPath || null,
      webUiCustomCssPath: path || null
    });
    showStatus('Web UI appearance saved. Reload the Web UI page to see changes.', 'info');
  } catch (error) {
    console.error('Failed to save Web UI appearance:', error);
    showStatus('Failed to save Web UI appearance', 'error');
  }
}

async function saveKeyboardShortcutSettings() {
  try {
    const validPresets = ['cmd+arrow', 'alt+arrow', 'cmd+shift+arrow'];
    const preset = keyboardShortcutPresetSelect ? keyboardShortcutPresetSelect.value : 'cmd+arrow';
    const defaultEnabled = keyboardShortcutsDefaultEnabledCheckbox ? keyboardShortcutsDefaultEnabledCheckbox.checked : false;
    await window.electronAPI.savePreferences({
      keyboardShortcutPreset: validPresets.includes(preset) ? preset : 'cmd+arrow',
      keyboardShortcutsDefaultEnabled: defaultEnabled
    });
    showStatus('Keyboard shortcut settings saved.', 'info');
  } catch (error) {
    console.error('Failed to save keyboard shortcut settings:', error);
    showStatus('Failed to save keyboard shortcut settings', 'error');
  }
}

// Save primary/backup preferences
async function savePrimaryBackupPreferences() {
  try {
    let mode = 'standalone';
    if (modePrimary.checked) {
      mode = 'primary';
    } else if (modeBackup.checked) {
      mode = 'backup';
    }
    
    const backupPort = parseInt(backupPortInput.value, 10);
    
    // Validate backup port
    if (mode === 'primary' && (isNaN(backupPort) || backupPort < 1024 || backupPort > 65535)) {
      showStatus('Backup port must be between 1024 and 65535', 'error');
      return;
    }

    const prefs = { primaryBackupMode: mode };
    if (mode === 'primary') {
      prefs.backupPort = backupPort;
      prefs.backupIps = getBackupIpsFromUi();
    }
    
    const saveResult = await window.electronAPI.savePreferences(prefs);

    // Restart backup status polling if needed
    if (mode === 'primary') {
      startBackupStatusPolling();
    } else {
      stopBackupStatusPolling();
    }

    if (saveResult && Array.isArray(saveResult.removedSelfReferentialBackupIps) && saveResult.removedSelfReferentialBackupIps.length > 0) {
      showStatus(
        `Saved. Removed this computer from backups: ${saveResult.removedSelfReferentialBackupIps.join(', ')}`,
        'error'
      );
      const refreshed = await window.electronAPI.getPreferences();
      renderBackupIpList(Array.isArray(refreshed.backupIps) ? refreshed.backupIps : []);
      refreshBackupStatusBadges();
    } else {
      showStatus('Primary/Backup configuration saved', 'info');
    }
  } catch (error) {
    console.error('Failed to save primary/backup preferences:', error);
    showStatus('Failed to save primary/backup preferences', 'error');
  }
}

// Backup status polling
let backupStatusInterval = null;

function startBackupStatusPolling() {
  stopBackupStatusPolling();
  
  // Poll immediately, then every 5 seconds
  updateBackupStatus();
  backupStatusInterval = setInterval(updateBackupStatus, 5000);
}

function stopBackupStatusPolling() {
  if (backupStatusInterval) {
    clearInterval(backupStatusInterval);
    backupStatusInterval = null;
  }
}

async function updateBackupStatus() {
  try {
    const preferences = await window.electronAPI.getPreferences();
    const apiPort = preferences.apiPort || 9595;
    
    const response = await fetch(`http://127.0.0.1:${apiPort}/api/backup-status`);
    if (!response.ok) {
      throw new Error('Failed to fetch backup status');
    }
    const data = await response.json();

    // Normalize into { ip -> status }
    backupStatusByIp = {};
    if (data && Array.isArray(data.backups)) {
      data.backups.forEach((b) => {
        const ip = String(b?.ip || '').trim();
        if (!ip) return;
        backupStatusByIp[ip] = b?.status || null;
      });
    }
    refreshBackupStatusBadges();
  } catch (error) {
    console.error('Failed to update backup status:', error);
  }
}

// Preset Presentations Functions
function getPresetUrlsFromUi() {
  if (!presetList) return [];
  const inputs = presetList.querySelectorAll('input[data-preset-url="true"]');
  return Array.from(inputs).map(inp => (inp.value || '').trim()).filter(Boolean);
}

function addPresetRow(initialValue = '') {
  if (!presetList) return;
  const row = document.createElement('div');
  row.setAttribute('data-preset-row', 'true');
  row.style.display = 'flex';
  row.style.gap = '10px';
  row.style.alignItems = 'center';

  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'input-field';
  input.placeholder = 'https://docs.google.com/presentation/d/...';
  input.value = initialValue || '';
  input.setAttribute('data-preset-url', 'true');
  input.style.flex = '1';

  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.className = 'btn btn-secondary';
  removeBtn.textContent = 'Remove';
  removeBtn.style.padding = '8px 10px';
  removeBtn.style.minWidth = '88px';
  removeBtn.addEventListener('click', () => {
    const rows = presetList.querySelectorAll('[data-preset-row="true"]');
    if (rows.length <= 1) {
      input.value = '';
      return;
    }
    row.remove();
  });

  row.appendChild(input);
  row.appendChild(removeBtn);
  presetList.appendChild(row);
}

function renderPresetList(urls = []) {
  if (!presetList) return;
  presetList.innerHTML = '';
  const normalized = Array.isArray(urls) ? urls : [];
  if (normalized.length === 0) {
    addPresetRow('');
    return;
  }
  normalized.forEach(u => addPresetRow(u));
}

async function loadPresets() {
  try {
    const preferences = await window.electronAPI.getPreferences();
    const urls = Array.isArray(preferences.presetUrls) ? preferences.presetUrls : [];
    if (urls.length === 0) {
      const legacy = [preferences.presentation1, preferences.presentation2, preferences.presentation3].filter(Boolean);
      renderPresetList(legacy.length ? legacy : ['']);
    } else {
      renderPresetList(urls);
    }
  } catch (error) {
    console.error('Failed to load presets:', error);
  }
}

async function savePresets() {
  try {
    const urls = getPresetUrlsFromUi();
    await window.electronAPI.savePreferences({ presetUrls: urls });
    showStatus('Presets saved successfully', 'info');
  } catch (error) {
    console.error('Failed to save presets:', error);
    showStatus('Failed to save presets', 'error');
  }
}

// Stagetimer Settings Functions
async function loadStagetimerSettings() {
  try {
    const preferences = await window.electronAPI.getPreferences();
    const apiPort = preferences.apiPort || 9595;
    
    const response = await fetch(`http://127.0.0.1:${apiPort}/api/stagetimer-settings`);
    if (!response.ok) {
      throw new Error('Failed to fetch stagetimer settings');
    }
    const data = await response.json();
    
    stagetimerRoomIdInput.value = data.roomId || '';
    stagetimerApiKeyInput.value = data.apiKey || '';
    stagetimerEnabledCheckbox.checked = data.enabled !== false;
    stagetimerVisibleCheckbox.checked = data.visible !== false && data.visible !== undefined ? data.visible : true;
  } catch (error) {
    console.error('Failed to load stagetimer settings:', error);
    // If API call fails, try loading from preferences directly
    const preferences = await window.electronAPI.getPreferences();
    stagetimerRoomIdInput.value = preferences.stagetimerRoomId || '';
    stagetimerApiKeyInput.value = preferences.stagetimerApiKey || '';
    stagetimerEnabledCheckbox.checked = preferences.stagetimerEnabled !== false;
    stagetimerVisibleCheckbox.checked = preferences.stagetimerVisible !== false && preferences.stagetimerVisible !== undefined ? preferences.stagetimerVisible : true;
  }
}

async function saveStagetimerSettings() {
  try {
    const preferences = await window.electronAPI.getPreferences();
    const apiPort = preferences.apiPort || 9595;
    
    const settings = {
      roomId: stagetimerRoomIdInput.value.trim(),
      apiKey: stagetimerApiKeyInput.value.trim(),
      enabled: stagetimerEnabledCheckbox.checked,
      visible: stagetimerVisibleCheckbox.checked
    };
    
    const response = await fetch(`http://127.0.0.1:${apiPort}/api/stagetimer-settings`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(settings)
    });
    
    if (!response.ok) {
      throw new Error('Failed to save stagetimer settings');
    }
    
    const result = await response.json();
    if (result.success) {
      showStatus('Stagetimer settings saved successfully', 'info');
    } else {
      showStatus('Failed to save stagetimer settings: ' + (result.error || 'Unknown error'), 'error');
    }
  } catch (error) {
    console.error('Failed to save stagetimer settings:', error);
    showStatus('Failed to save stagetimer settings: ' + error.message, 'error');
  }
}

// Show status message
function showStatus(message, type = 'info') {
  statusMessage.textContent = message;
  statusMessage.className = `status-message show ${type}`;
  
  setTimeout(() => {
    statusMessage.classList.remove('show');
  }, 4000);
}

// Update auth status UI
function updateAuthStatus(signedIn) {
  isSignedIn = signedIn;
  
  if (signedIn) {
    signinBtn.innerHTML = `
      <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path>
        <polyline points="16 17 21 12 16 7"></polyline>
        <line x1="21" y1="12" x2="9" y2="12"></line>
      </svg>
      Sign Out
    `;
    signinBtn.disabled = false;
    signinBtn.classList.remove('btn-google');
    signinBtn.classList.add('btn-signout');
  } else {
    signinBtn.innerHTML = `
      <svg class="btn-icon" viewBox="0 0 24 24">
        <path fill="currentColor" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>
        <path fill="currentColor" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>
        <path fill="currentColor" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/>
        <path fill="currentColor" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/>
      </svg>
      Sign in with Google
    `;
    signinBtn.disabled = false;
    signinBtn.classList.add('btn-google');
    signinBtn.classList.remove('btn-signout');
  }
  updateStatusBar();
}

// Sign in/out with Google
signinBtn.addEventListener('click', async () => {
  if (isSignedIn) {
    // Sign out
    signinBtn.disabled = true;
    signinBtn.textContent = 'Signing out...';
    
    try {
      const result = await window.electronAPI.googleSignOut();
      if (result.success) {
        updateAuthStatus(false);
        showStatus('Successfully signed out', 'success');
      }
    } catch (error) {
      showStatus('Failed to sign out', 'error');
      signinBtn.disabled = false;
    }
  } else {
    // Sign in
    signinBtn.disabled = true;
    signinBtn.textContent = 'Opening sign in...';
    
    try {
      const result = await window.electronAPI.googleSignIn();
      if (result.success) {
        updateAuthStatus(true);
        showStatus('Successfully signed in to Google!', 'success');
      }
    } catch (error) {
      showStatus(error.message || 'Sign in was cancelled', 'error');
      updateAuthStatus(false);
    }
  }
});

// Open test presentation
testBtn.addEventListener('click', async () => {
  testBtn.disabled = true;
  const originalText = testBtn.innerHTML;
  testBtn.innerHTML = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>Opening...';
  
  try {
    const result = await window.electronAPI.openTestPresentation();
    if (result.success) {
      showStatus('Test presentation opened!', 'success');
    }
  } catch (error) {
    showStatus('Failed to open test presentation', 'error');
  } finally {
    setTimeout(() => {
      testBtn.disabled = false;
      testBtn.innerHTML = originalText;
    }, 1000);
  }
});

// Check sign-in status on load
async function checkSignInStatus() {
  try {
    const status = await window.electronAPI.checkSignInStatus();
    if (status.signedIn) {
      updateAuthStatus(true);
    }
  } catch (error) {
    console.error('Failed to check sign-in status:', error);
  }
}

// Load and display build number
async function displayBuildNumber() {
  try {
    const buildInfo = await window.electronAPI.getBuildInfo();
    const version = buildInfo.version || '2.2.1';
    const buildNumber = buildInfo.buildNumber || '72';
    const versionString = `v${version}.${buildNumber}`;
    
    const buildNumberEl = document.getElementById('build-number');
    if (buildNumberEl) {
      const chrome = buildInfo.chrome ? ` · Chromium ${buildInfo.chrome}` : '';
      buildNumberEl.textContent = versionString + chrome;
    }
  } catch (error) {
    console.error('Failed to load build number:', error);
    const buildNumberEl = document.getElementById('build-number');
    if (buildNumberEl) {
      buildNumberEl.textContent = 'v2.2.1.72';
    }
  }
}

// Initialize on load
initDisplays();
checkSignInStatus();
displayBuildNumber();
initDebugLogs();