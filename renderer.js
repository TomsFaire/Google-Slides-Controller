// DOM Elements
const presentationDisplay = document.getElementById('presentation-display');
const notesDisplay = document.getElementById('notes-display');
const apiPortInput = document.getElementById('api-port');
const webUiPortInput = document.getElementById('web-ui-port');
const signinBtn = document.getElementById('signin-btn');
const testBtn = document.getElementById('test-btn');
const statusMessage = document.getElementById('status-message');

let isSignedIn = false;

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
    
    // Restore saved preferences or use defaults
    if (preferences.presentationDisplayId) {
      presentationDisplay.value = preferences.presentationDisplayId;
    }
    
    if (preferences.notesDisplayId) {
      notesDisplay.value = preferences.notesDisplayId;
    } else if (displays.length > 1 && !preferences.notesDisplayId) {
      // Select different displays by default if available and no preference saved
      notesDisplay.selectedIndex = 1;
    }
    
    // Restore machine name
    const machineNameInput = document.getElementById('machine-name');
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
    
    // Save preferences when selections change
    presentationDisplay.addEventListener('change', saveMonitorPreferences);
    notesDisplay.addEventListener('change', saveMonitorPreferences);
    machineNameInput.addEventListener('change', saveMachineName);
    apiPortInput.addEventListener('change', savePortPreferences);
    webUiPortInput.addEventListener('change', savePortPreferences);
    
    // Load and display network info
    await updateNetworkInfo();
    
  } catch (error) {
    showStatus('Failed to load displays', 'error');
  }
}

// Update network information display
async function updateNetworkInfo() {
  try {
    const networkInfo = await window.electronAPI.getNetworkInfo();
    const preferences = await window.electronAPI.getPreferences();
    
    const apiPort = preferences.apiPort || 9595;
    const webUiPort = preferences.webUiPort || 80;
    
    // Display API URLs
    const apiUrlsDiv = document.getElementById('api-urls');
    apiUrlsDiv.innerHTML = '';
    
    if (networkInfo.length === 0) {
      apiUrlsDiv.innerHTML = '<div class="url-item">No network interfaces found</div>';
    } else {
      networkInfo.forEach(ip => {
        const urlItem = document.createElement('div');
        urlItem.className = 'url-item' + (ip.internal ? ' internal' : '');
        urlItem.textContent = `http://${ip.address}:${apiPort}`;
        if (ip.internal) {
          urlItem.title = 'Localhost/internal interface';
        }
        apiUrlsDiv.appendChild(urlItem);
      });
    }
    
    // Display Web UI URLs
    const webUiUrlsDiv = document.getElementById('web-ui-urls');
    webUiUrlsDiv.innerHTML = '';
    
    if (networkInfo.length === 0) {
      webUiUrlsDiv.innerHTML = '<div class="url-item">No network interfaces found</div>';
    } else {
      networkInfo.forEach(ip => {
        const urlItem = document.createElement('div');
        urlItem.className = 'url-item' + (ip.internal ? ' internal' : '');
        urlItem.textContent = `http://${ip.address}:${webUiPort}`;
        if (ip.internal) {
          urlItem.title = 'Localhost/internal interface';
        }
        webUiUrlsDiv.appendChild(urlItem);
      });
    }
  } catch (error) {
    console.error('Failed to load network info:', error);
    document.getElementById('api-urls').innerHTML = '<div class="url-item">Error loading network info</div>';
    document.getElementById('web-ui-urls').innerHTML = '<div class="url-item">Error loading network info</div>';
  }
}

// Save monitor preferences
async function saveMonitorPreferences() {
  try {
    await window.electronAPI.savePreferences({
      presentationDisplayId: presentationDisplay.value,
      notesDisplayId: notesDisplay.value
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
    } catch (error) {
      console.error('Failed to save machine name:', error);
      showStatus('Failed to save machine name', 'error');
    }
  }

  // Save port preferences
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
    
    await window.electronAPI.savePreferences({
      apiPort: apiPort,
      webUiPort: webUiPort
    });
    
    // Update network info display with new ports
    await updateNetworkInfo();
    
    showStatus('Ports saved. Please restart the app for changes to take effect.', 'info');
  } catch (error) {
    console.error('Failed to save port preferences:', error);
    showStatus('Failed to save port preferences', 'error');
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

// Initialize on load
initDisplays();
checkSignInStatus();
