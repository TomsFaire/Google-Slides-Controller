const { app, BrowserWindow, ipcMain, screen, session, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const http = require('http');
const https = require('https');
const os = require('os');

let mainWindow;
let presentationWindow = null;
let notesWindow = null;
let currentSlide = null; // best-effort: we track on our next/prev; DOM can override when notes window has aria-posinset/aria-setsize
let lastPresentationUrl = null; // Store the last-opened presentation URL for reload functionality

function toPresentUrl(inputUrl) {
  try {
    const u = new URL(inputUrl);

    // Extract deck id from any /presentation/d/<ID>/... path
    const m = u.pathname.match(/\/presentation\/d\/([^/]+)/);
    if (!m) return inputUrl;

    const id = m[1];

    // Go straight to slideshow mode (avoids platform-specific present hotkeys)
    return `https://docs.google.com/presentation/d/${id}/present`;
  } catch (e) {
    return inputUrl;
  }
}


// Use a persistent session for Google authentication
const GOOGLE_SESSION_PARTITION = 'persist:google';

// Function to set speaker notes window to fullscreen
function setSpeakerNotesFullscreen(window) {
  if (!window || window.isDestroyed()) return;
  
  // Hide the window initially to prevent seeing the resize
  window.hide();
  
  // Wait for page to fully load before showing
  const showFullscreen = () => {
    if (window.isDestroyed()) return;
    
    const display = screen.getDisplayMatching(window.getBounds());
    window.setBounds(display.bounds);
    if (process.platform === 'darwin') {
      window.setFullScreen(true);
    }
    window.show();
    console.log('[Notes] Set speaker notes window to fullscreen');
  };
  
  // Wait for page to finish loading, then wait a bit more for layout to stabilize
  window.webContents.once('did-finish-load', () => {
    // Wait a bit longer for images and layout to settle
    setTimeout(showFullscreen, 1500);
  });
  
  // Fallback: if did-finish-load already fired, use dom-ready
  if (window.webContents.isLoading() === false) {
    setTimeout(showFullscreen, 1500);
  } else {
    window.webContents.once('dom-ready', () => {
      setTimeout(showFullscreen, 1500);
    });
  }
}

function getGoogleSession() {
  return session.fromPartition(GOOGLE_SESSION_PARTITION);
}

// Get preferences file path
function getPreferencesPath() {
  return path.join(app.getPath('userData'), 'preferences.json');
}

// Load preferences
function loadPreferences() {
  try {
    const prefsPath = getPreferencesPath();
    console.log('[Preferences] Loading from:', prefsPath);
    
    if (fs.existsSync(prefsPath)) {
      const data = fs.readFileSync(prefsPath, 'utf8');
      const prefs = JSON.parse(data);
      console.log('[Preferences] Loaded preferences:', JSON.stringify(prefs));
      return prefs;
    } else {
      console.log('[Preferences] File does not exist, returning empty object');
    }
  } catch (error) {
    console.error('[Preferences] Error loading preferences:', error);
    console.error('[Preferences] Error details:', {
      message: error.message,
      code: error.code,
      path: getPreferencesPath()
    });
  }
  return {};
}

// Save preferences
function savePreferences(prefs) {
  try {
    const prefsPath = getPreferencesPath();
    console.log('[Preferences] Saving to:', prefsPath);
    console.log('[Preferences] Data to save:', JSON.stringify(prefs, null, 2));
    
    // Ensure directory exists
    const prefsDir = path.dirname(prefsPath);
    if (!fs.existsSync(prefsDir)) {
      console.log('[Preferences] Creating directory:', prefsDir);
      fs.mkdirSync(prefsDir, { recursive: true });
    }
    
    fs.writeFileSync(prefsPath, JSON.stringify(prefs, null, 2), 'utf8');
    console.log('[Preferences] Successfully saved preferences');
    
    // Verify it was written
    if (fs.existsSync(prefsPath)) {
      const stats = fs.statSync(prefsPath);
      console.log('[Preferences] File verified - size:', stats.size, 'bytes');
    } else {
      console.error('[Preferences] ERROR: File was not created after write!');
    }
  } catch (error) {
    console.error('[Preferences] Error saving preferences:', error);
    console.error('[Preferences] Error details:', {
      message: error.message,
      code: error.code,
      path: getPreferencesPath(),
      stack: error.stack
    });
    throw error; // Re-throw so caller can handle it
  }
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 900,
    minWidth: 600,
    minHeight: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    },
    autoHideMenuBar: true,
    resizable: true,
    center: true
  });

  mainWindow.loadFile('index.html');
  
  // Open DevTools for main window to see logs
  // mainWindow.webContents.openDevTools();
}

// Get all available displays
ipcMain.handle('get-displays', async () => {
  const displays = screen.getAllDisplays();
  return displays.map((display, index) => ({
    id: display.id,
    label: `Monitor ${index + 1} (${display.bounds.width}x${display.bounds.height})`,
    bounds: display.bounds,
    primary: display.bounds.x === 0 && display.bounds.y === 0
  }));
});

// Get saved preferences
ipcMain.handle('get-preferences', async () => {
  return loadPreferences();
});

// Get build info (version and build number)
ipcMain.handle('get-build-info', async () => {
  try {
    const packageJsonPath = path.join(__dirname, 'package.json');
    const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
      return {
        version: packageJson.version || '1.4.5',
        buildNumber: packageJson.buildNumber || '24'
      };
  } catch (error) {
    console.error('[Build Info] Error loading build info:', error);
    return {
      version: '1.4.5',
      buildNumber: '24'
    };
  }
});

// Get network interfaces and IP addresses
ipcMain.handle('get-network-info', async () => {
  const interfaces = os.networkInterfaces();
  const ipAddresses = [];
  
  // Get all IPv4 addresses (excluding internal/loopback, but including localhost)
  Object.keys(interfaces).forEach((ifaceName) => {
    interfaces[ifaceName].forEach((iface) => {
      // Include IPv4 addresses (both internal and external)
      if (iface.family === 'IPv4') {
        ipAddresses.push({
          address: iface.address,
          internal: iface.internal,
          interface: ifaceName
        });
      }
    });
  });
  
  // Sort: non-internal first, then by interface name
  ipAddresses.sort((a, b) => {
    if (a.internal !== b.internal) {
      return a.internal ? 1 : -1;
    }
    return a.interface.localeCompare(b.interface);
  });
  
  return ipAddresses;
});

// Save preferences
ipcMain.handle('save-preferences', async (event, prefs) => {
  const currentPrefs = loadPreferences();
  const mergedPrefs = { ...currentPrefs, ...prefs };
  savePreferences(mergedPrefs);
  return { success: true };
});

// Sign in with Google
ipcMain.handle('google-signin', async () => {
  const googleSession = getGoogleSession();
  
  const authWindow = new BrowserWindow({
    width: 500,
    height: 600,
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      partition: GOOGLE_SESSION_PARTITION
    }
  });

  // Navigate to Google Sign In
  authWindow.loadURL('https://accounts.google.com/signin');

  authWindow.once('ready-to-show', () => {
    authWindow.show();
  });

  // Listen for successful authentication
  return new Promise((resolve, reject) => {
    let resolved = false;
    
    authWindow.webContents.on('did-navigate', (event, url) => {
      // Check if we've successfully signed in (redirected to myaccount or other Google service)
      if (url.includes('myaccount.google.com') || url.includes('accounts.google.com/ServiceLogin')) {
        if (!resolved) {
          resolved = true;
          authWindow.close();
          resolve({ success: true, message: 'Successfully signed in to Google' });
        }
      }
    });

    authWindow.on('closed', () => {
      if (!resolved) {
        resolved = true;
        reject({ success: false, message: 'Authentication window closed' });
      }
    });
  });
});

// Check if user is already signed in
ipcMain.handle('check-signin-status', async () => {
  try {
    const googleSession = getGoogleSession();
    const cookies = await googleSession.cookies.get({ domain: '.google.com' });
    
    // Check if we have Google authentication cookies
    const hasAuthCookies = cookies.some(cookie => 
      cookie.name === 'SID' || cookie.name === 'HSID' || cookie.name === 'SSID'
    );
    
    let userEmail = null;
    let userName = null;
    
    if (hasAuthCookies) {
      // Try to get user email from cookies
      const emailCookie = cookies.find(cookie => 
        cookie.name === 'Email' || cookie.name === 'email' || cookie.domain.includes('google.com')
      );
      if (emailCookie && emailCookie.value && emailCookie.value.includes('@')) {
        userEmail = emailCookie.value;
      }
      
      // Try to get user name from cookies
      const nameCookie = cookies.find(cookie => 
        cookie.name === 'Name' || cookie.name === 'name'
      );
      if (nameCookie) {
        userName = nameCookie.value;
      }
      
      // If we don't have email, try to fetch from Google account page
      if (!userEmail) {
        try {
          const tempWindow = new BrowserWindow({
            show: false,
            webPreferences: {
              nodeIntegration: false,
              contextIsolation: true,
              partition: GOOGLE_SESSION_PARTITION
            }
          });
          
          await tempWindow.loadURL('https://myaccount.google.com/');
          await new Promise(resolve => setTimeout(resolve, 2000));
          
          const userInfo = await tempWindow.webContents.executeJavaScript(`
            (function() {
              try {
                var email = null;
                var name = null;
                
                // Look for email in page
                var emailEl = document.querySelector('[data-email]') || 
                             document.querySelector('input[type="email"][value]');
                if (emailEl) {
                  email = emailEl.getAttribute('data-email') || emailEl.value;
                }
                
                // Look for name
                var nameEl = document.querySelector('[data-name]') ||
                            document.querySelector('h1');
                if (nameEl) {
                  name = nameEl.getAttribute('data-name') || nameEl.textContent.trim();
                }
                
                // Try to extract from page title
                if (!email) {
                  var title = document.title;
                  var emailMatch = title.match(/([\\w.-]+@[\\w.-]+\\.[\\w.-]+)/);
                  if (emailMatch) email = emailMatch[1];
                }
                
                return { email: email || null, name: name || null };
              } catch (e) {
                return { email: null, name: null };
              }
            })()
          `);
          
          if (userInfo.email) userEmail = userInfo.email;
          if (userInfo.name) userName = userInfo.name;
          
          tempWindow.close();
        } catch (error) {
          console.error('Error fetching user info:', error);
        }
      }
    }
    
    return { 
      signedIn: hasAuthCookies,
      userEmail: userEmail || null,
      userName: userName || null
    };
  } catch (error) {
    console.error('Error checking sign-in status:', error);
    return { signedIn: false, userEmail: null, userName: null };
  }
});

// Sign out from Google
ipcMain.handle('google-signout', async () => {
  const googleSession = getGoogleSession();
  
  // Clear all cookies and storage data for the Google session
  await googleSession.clearStorageData({
    storages: ['cookies', 'localstorage', 'sessionstorage', 'cachestorage']
  });
  
  return { success: true, message: 'Successfully signed out' };
});

// Open test presentation
ipcMain.handle('open-test-presentation', async () => {
  const testUrl = 'https://docs.google.com/presentation/d/1rc9BSX-0TrU7c5LGeLDRyH3zRN89-uDuXEEqOpcnLVg/edit';
  
  // Load preferences to get selected displays
  const prefs = loadPreferences();
  console.log('[Test] Loaded preferences:', prefs);
  
  const displays = screen.getAllDisplays();
  console.log('[Test] All available displays:');
  displays.forEach((display, index) => {
    console.log(`  Display ${index + 1} - ID: ${display.id}, Bounds: ${JSON.stringify(display.bounds)}`);
  });
  
  // Convert IDs to numbers for comparison (they might be saved as strings)
  const presentationDisplayId = Number(prefs.presentationDisplayId);
  const notesDisplayId = Number(prefs.notesDisplayId);
  
  const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
  const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];
  
  console.log('[Test] Selected presentation display ID:', prefs.presentationDisplayId, '(converted to:', presentationDisplayId, ')');
  console.log('[Test] Resolved presentation display:', presentationDisplay.id, 'Bounds:', presentationDisplay.bounds);
  console.log('[Test] Selected notes display ID:', prefs.notesDisplayId, '(converted to:', notesDisplayId, ')');
  console.log('[Test] Resolved notes display:', notesDisplay.id, 'Bounds:', notesDisplay.bounds);
  
  if (!presentationWindow) {
    presentationWindow = new BrowserWindow({
      x: presentationDisplay.bounds.x,
      y: presentationDisplay.bounds.y,
      width: presentationDisplay.bounds.width,
      height: presentationDisplay.bounds.height,
      fullscreen: true,
      frame: false,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        partition: GOOGLE_SESSION_PARTITION
      }
    });

    presentationWindow.on('closed', () => {
      presentationWindow = null;
      currentSlide = null;
    });
    
    // Listen for Escape key to close both windows
    presentationWindow.webContents.on('before-input-event', (event, input) => {
      if (input.key === 'Escape' && input.type === 'keyDown') {
        console.log('[Test] Escape pressed, closing presentation and notes windows');
        event.preventDefault(); // Prevent Google Slides from handling Escape
        if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
        if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
      }
    });

    // Handle the speaker notes popup window
    presentationWindow.webContents.setWindowOpenHandler((details) => {
      console.log('[Test] Window open intercepted:', details.url);
      console.log('[Test] Frame name:', details.frameName);
      console.log('[Test] Features:', details.features);
      
      // Allow Google Slides to open the speaker notes window
      // Use default size from Google Slides (no size/position override)
      const windowOptions = {
        frame: false,
        webPreferences: {
          nodeIntegration: false,
          contextIsolation: true,
          partition: GOOGLE_SESSION_PARTITION
        }
      };
      
      return {
        action: 'allow',
        overrideBrowserWindowOptions: windowOptions
      };
    });
    
    // Listen for new windows being created (this will be the notes window)
    const testWindowCreatedListener = (event, window) => {
      if (window !== presentationWindow && window !== mainWindow) {
        console.log('[Test] Notes window created');
        console.log('[Test] Presentation display ID:', presentationDisplay.id);
        console.log('[Test] Notes display ID:', notesDisplay.id);
        notesWindow = window;
        
        const initialBounds = window.getBounds();
        console.log('[Test] Initial window bounds:', initialBounds);
        
        // Add Escape key handler to notes window as well
        window.webContents.on('before-input-event', (event, input) => {
          if (input.key === 'Escape' && input.type === 'keyDown') {
            console.log('[Test] Escape pressed in notes window, closing all windows');
            event.preventDefault();
            if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
            if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
          }
        });
        
        // Set speaker notes window to fullscreen when it loads
        window.webContents.once('did-finish-load', () => {
          setSpeakerNotesFullscreen(window);
        });
        
        // Also try when DOM is ready (in case did-finish-load fires too early)
        window.webContents.once('dom-ready', () => {
          setTimeout(() => {
            setSpeakerNotesFullscreen(window);
          }, 500);
        });
        
        app.removeListener('browser-window-created', testWindowCreatedListener);
      }
    };
    app.on('browser-window-created', testWindowCreatedListener);
  }

  lastPresentationUrl = testUrl; // Store for reload
  currentSlide = 1;
  presentationWindow.loadURL(testUrl);
  presentationWindow.show();
  
  console.log('[Test] Window opened, loading URL...');
  
  // Set up navigation listener
  const navigationListener = async (event, url) => {
    console.log('[Test] Navigated to:', url);
    
    // Just log navigation, don't auto-launch notes
    console.log('[Test] Navigated to:', url);
  };
  
  presentationWindow.webContents.on('did-navigate', navigationListener);
  
  // Listen for page load, then immediately trigger presentation mode
  presentationWindow.webContents.once('did-finish-load', async () => {
    if (!presentationWindow || presentationWindow.isDestroyed()) return;
    
    const currentUrl = presentationWindow.webContents.getURL();
    console.log('[Test] Page loaded:', currentUrl);
    
    // Small delay to ensure page is fully interactive
    await new Promise(resolve => setTimeout(resolve, 200));
    
    if (!presentationWindow || presentationWindow.isDestroyed()) return;
    
    console.log('[Test] Triggering Ctrl+Shift+F5 to enter presentation mode...');
    
    try {
      // Focus the window first to ensure it receives the keyboard events
      presentationWindow.focus();
      await new Promise(resolve => setTimeout(resolve, 50));
      
      // Send real keyboard input events
      presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
      presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
      
      console.log('[Test] Ctrl+Shift+F5 sent via sendInputEvent');
    } catch (error) {
      console.error('[Test] Error sending Ctrl+Shift+F5:', error);
    }
    
    // No auto-launch of speaker notes - user must call open-speaker-notes separately
  });
  
  return { success: true };
});

// Open presentation on specific monitor
ipcMain.handle('open-presentation', async (event, { url, presentationDisplayId, notesDisplayId }) => {
  const displays = screen.getAllDisplays();
  console.log('[Multi-Monitor] All available displays:');
  displays.forEach((display, index) => {
    console.log(`  Display ${index + 1} - ID: ${display.id}, Bounds: ${JSON.stringify(display.bounds)}`);
  });
  
  // Convert IDs to numbers for comparison (they might be passed as strings)
  const presentationDisplayIdNum = Number(presentationDisplayId);
  const notesDisplayIdNum = Number(notesDisplayId);
  
  const presentationDisplay = displays.find(d => d.id === presentationDisplayIdNum);
  const notesDisplay = displays.find(d => d.id === notesDisplayIdNum);

  console.log('[Multi-Monitor] Selected presentation display ID:', presentationDisplayId, '(converted to:', presentationDisplayIdNum, ')');
  console.log('[Multi-Monitor] Resolved presentation display:', presentationDisplay ? presentationDisplay.id : 'NOT FOUND', 'Bounds:', presentationDisplay ? presentationDisplay.bounds : 'N/A');
  console.log('[Multi-Monitor] Selected notes display ID:', notesDisplayId, '(converted to:', notesDisplayIdNum, ')');
  console.log('[Multi-Monitor] Resolved notes display:', notesDisplay ? notesDisplay.id : 'NOT FOUND', 'Bounds:', notesDisplay ? notesDisplay.bounds : 'N/A');

  if (!presentationDisplay) {
    return { success: false, message: 'Invalid presentation display' };
  }

  // Close existing windows if any
  if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
  if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
  currentSlide = null;

  // Open presentation window
  presentationWindow = new BrowserWindow({
    x: presentationDisplay.bounds.x,
    y: presentationDisplay.bounds.y,
    width: presentationDisplay.bounds.width,
    height: presentationDisplay.bounds.height,
    fullscreen: true,
    frame: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      partition: GOOGLE_SESSION_PARTITION
    }
  });

  // Handle the speaker notes popup window
  presentationWindow.webContents.setWindowOpenHandler((details) => {
    console.log('[Multi-Monitor] Window open intercepted:', details.url);
    console.log('[Multi-Monitor] Frame name:', details.frameName);
    console.log('[Multi-Monitor] Features:', details.features);
    
    // Allow Google Slides to open the speaker notes window
    // Use default size from Google Slides (no size/position override)
    const windowOptions = {
      frame: false,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        partition: GOOGLE_SESSION_PARTITION
      }
    };
    
    return {
      action: 'allow',
      overrideBrowserWindowOptions: windowOptions
    };
  });
  
  // Listen for new windows being created (this will be the notes window)
  const windowCreatedListener = (event, window) => {
    // Check if this is not the presentation window or main window
    if (window !== presentationWindow && window !== mainWindow) {
      console.log('[Multi-Monitor] Notes window created');
      console.log('[Multi-Monitor] Presentation display ID:', presentationDisplayIdNum);
      console.log('[Multi-Monitor] Notes display ID:', notesDisplayIdNum);
      console.log('[Multi-Monitor] Notes display object:', notesDisplay);
      notesWindow = window;
      
      // Get initial window bounds
      const initialBounds = window.getBounds();
      console.log('[Multi-Monitor] Initial window bounds:', initialBounds);
      
      // Add Escape key handler to notes window as well
      window.webContents.on('before-input-event', (event, input) => {
        if (input.key === 'Escape' && input.type === 'keyDown') {
          console.log('[Multi-Monitor] Escape pressed in notes window, closing all windows');
          event.preventDefault();
          if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
          if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
        }
      });
      
      // Set speaker notes window to fullscreen when it loads
      window.webContents.once('did-finish-load', () => {
        setSpeakerNotesFullscreen(window);
      });
      
      // Also try when DOM is ready (in case did-finish-load fires too early)
      window.webContents.once('dom-ready', () => {
        setTimeout(() => {
          setSpeakerNotesFullscreen(window);
        }, 500);
      });
      
      // Remove listener after notes window is created
      app.removeListener('browser-window-created', windowCreatedListener);
    }
  };
  app.on('browser-window-created', windowCreatedListener);

  // Load presentation URL
  lastPresentationUrl = url; // Store for reload
  currentSlide = 1;
  presentationWindow.loadURL(url);

  console.log('[Multi-Monitor] Window opened, loading URL...');

  // Listen for all page loads
  // Set up navigation listener to detect presentation mode activation
  const navigationListener = async (event, url) => {
    console.log('[Multi-Monitor] Navigated to:', url);
    // Just log navigation, don't auto-launch notes
  };
  
  presentationWindow.webContents.on('did-navigate', navigationListener);
  
  // Listen for page load, then immediately trigger presentation mode
  presentationWindow.webContents.once('did-finish-load', async () => {
    if (!presentationWindow || presentationWindow.isDestroyed()) return;
    
    const currentUrl = presentationWindow.webContents.getURL();
    console.log('[Multi-Monitor] Page loaded:', currentUrl);
    
    // Small delay to ensure page is fully interactive
    await new Promise(resolve => setTimeout(resolve, 200));
    
    if (!presentationWindow || presentationWindow.isDestroyed()) return;
    
    console.log('[Multi-Monitor] Triggering Ctrl+Shift+F5 to enter presentation mode...');
    
    try {
      // Focus the window first to ensure it receives the keyboard events
      presentationWindow.focus();
      await new Promise(resolve => setTimeout(resolve, 50));
      
      // Send real keyboard input events
      presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
      presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
      
      console.log('[Multi-Monitor] Ctrl+Shift+F5 sent via sendInputEvent');
    } catch (error) {
      console.error('[Multi-Monitor] Error sending Ctrl+Shift+F5:', error);
    }
    
    // No auto-launch of speaker notes - user must call open-speaker-notes separately
  });

  presentationWindow.on('closed', () => {
    presentationWindow = null;
    currentSlide = null;
  });
  
  // Listen for Escape key to close both windows
  presentationWindow.webContents.on('before-input-event', (event, input) => {
    if (input.key === 'Escape' && input.type === 'keyDown') {
      console.log('[Multi-Monitor] Escape pressed, closing presentation and notes windows');
      event.preventDefault(); // Prevent Google Slides from handling Escape
      if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
      if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
    }
  });

  return { success: true };
});

// HTTP API for Bitfocus Companion integration
// Ports are configurable via preferences, defaults below
const DEFAULT_API_PORT = 9595;
const DEFAULT_WEB_UI_PORT = 80;
let httpServer;
let webUiServer;

function startHttpServer() {
  httpServer = http.createServer(async (req, res) => {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    
    if (req.method === 'OPTIONS') {
      res.writeHead(200);
      res.end();
      return;
    }
    
    // GET /api/status - Check if app is running and expose state for Companion variables/feedbacks
    if (req.method === 'GET' && req.url === '/api/status') {
      (async () => {
        // Get login state and user info
        let loginState = false;
        let loggedInUser = null;
        try {
          const googleSession = getGoogleSession();
          const cookies = await googleSession.cookies.get({ domain: '.google.com' });
          const hasAuthCookies = cookies.some(cookie => 
            cookie.name === 'SID' || cookie.name === 'HSID' || cookie.name === 'SSID'
          );
          loginState = hasAuthCookies;
          
          if (hasAuthCookies) {
            // Try to get user email from cookies
            const emailCookie = cookies.find(cookie => 
              cookie.name === 'Email' || cookie.name === 'email' || 
              (cookie.value && cookie.value.includes('@'))
            );
            if (emailCookie && emailCookie.value && emailCookie.value.includes('@')) {
              loggedInUser = emailCookie.value;
            } else {
              // Try to get from any cookie value that looks like an email
              const emailLikeCookie = cookies.find(cookie => 
                cookie.value && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cookie.value)
              );
              if (emailLikeCookie) {
                loggedInUser = emailLikeCookie.value;
              }
            }
          }
        } catch (error) {
          console.error('[API] Error checking login state:', error);
        }
        
        const state = {
          status: 'ok',
          version: '1.4.5',
          presentationOpen: !!(presentationWindow && !presentationWindow.isDestroyed()),
          notesOpen: !!(notesWindow && !notesWindow.isDestroyed()),
          currentSlide: currentSlide,
          totalSlides: null,
          presentationUrl: lastPresentationUrl || null,
          slideInfo: null,
          isFirstSlide: null,
          isLastSlide: null,
          nextSlide: null,
          previousSlide: null,
          presentationTitle: null,
          timerElapsed: null,
          loginState: loginState,
          loggedInUser: loggedInUser || null
        };
        
        // Get slide info and other data from notes window DOM
        if (notesWindow && !notesWindow.isDestroyed()) {
          try {
            const info = await notesWindow.webContents.executeJavaScript(`
              (function(){
                var result = {};
                
                // Get slide numbers from aria attributes
                var el = document.querySelector('[aria-posinset]');
                if (el) {
                  var cur = parseInt(el.getAttribute('aria-posinset'), 10);
                  var tot = parseInt(el.getAttribute('aria-setsize'), 10);
                  if (!isNaN(cur)) result.current = cur;
                  if (!isNaN(tot)) result.total = tot;
                }
                
                // Get presentation title from page title or DOM
                var titleEl = document.querySelector('title');
                if (titleEl) {
                  var titleText = titleEl.textContent;
                  // Extract title from "Presenter view - TITLE - Google Slides"
                  var match = titleText.match(/Presenter view - (.+?) - Google Slides/);
                  if (match) {
                    result.title = match[1];
                  } else {
                    result.title = titleText;
                  }
                }
                
                // Get timer value (look for timer display - usually shows "00:00:06" format)
                // Try to find elements containing time format
                var allText = document.body.innerText || document.body.textContent || '';
                var timeMatch = allText.match(/(\\d{1,2}:\\d{2}(?::\\d{2})?)/);
                if (timeMatch) {
                  result.timer = timeMatch[1];
                } else {
                  // Try specific timer elements
                  var timerEls = document.querySelectorAll('div, span');
                  for (var i = 0; i < timerEls.length; i++) {
                    var text = timerEls[i].textContent || timerEls[i].innerText || '';
                    var match = text.match(/^(\\d{1,2}:\\d{2}(?::\\d{2})?)$/);
                    if (match) {
                      result.timer = match[1];
                      break;
                    }
                  }
                }
                
                return result;
              })()
            `);
            
            if (info) {
              if (info.current != null) {
                state.currentSlide = info.current;
                // Calculate derived values
                if (info.total != null) {
                  state.totalSlides = info.total;
                  state.isFirstSlide = info.current === 1;
                  state.isLastSlide = info.current === info.total;
                  state.nextSlide = info.current < info.total ? info.current + 1 : null;
                  state.previousSlide = info.current > 1 ? info.current - 1 : null;
                  state.slideInfo = info.current + ' / ' + info.total;
                } else if (state.currentSlide !== null) {
                  // Use tracked currentSlide if DOM didn't provide total
                  state.isFirstSlide = state.currentSlide === 1;
                  state.nextSlide = state.currentSlide + 1;
                  state.previousSlide = state.currentSlide > 1 ? state.currentSlide - 1 : null;
                  if (state.totalSlides) {
                    state.isLastSlide = state.currentSlide === state.totalSlides;
                    state.slideInfo = state.currentSlide + ' / ' + state.totalSlides;
                  } else {
                    state.slideInfo = String(state.currentSlide);
                  }
                }
              }
              
              if (info.title) state.presentationTitle = info.title;
              if (info.timer) state.timerElapsed = info.timer;
            }
          } catch (e) { /* DOM not available or changed */ }
        }
        
        // Calculate derived values even if DOM didn't provide them
        if (state.currentSlide !== null && state.currentSlide !== undefined) {
          if (state.isFirstSlide === null) state.isFirstSlide = state.currentSlide === 1;
          if (state.nextSlide === null) state.nextSlide = state.currentSlide + 1;
          if (state.previousSlide === null) state.previousSlide = state.currentSlide > 1 ? state.currentSlide - 1 : null;
          if (state.slideInfo === null) {
            if (state.totalSlides) {
              state.slideInfo = state.currentSlide + ' / ' + state.totalSlides;
            } else {
              state.slideInfo = String(state.currentSlide);
            }
          }
          if (state.totalSlides && state.isLastSlide === null) {
            state.isLastSlide = state.currentSlide === state.totalSlides;
          }
        }
        
        // Get preferences for display IDs
        const prefs = loadPreferences();
        state.presentationDisplayId = prefs.presentationDisplayId || null;
        state.notesDisplayId = prefs.notesDisplayId || null;
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(state));
      })().catch(err => {
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: err.message }));
        }
      });
      return;
    }
    
    // POST /api/open-presentation - Open a presentation with URL
    if (req.method === 'POST' && req.url === '/api/open-presentation') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      
      req.on('end', async () => {
        try {
          const data = JSON.parse(body);
          const { url } = data;
          
          if (!url) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'URL is required' }));
            return;
          }
          
          console.log('[API] Opening presentation:', url);
          
          // Close any existing presentation windows
          try {
            if (notesWindow && !notesWindow.isDestroyed()) {
              console.log('[API] Closing existing notes window');
              notesWindow.removeAllListeners('closed');
              notesWindow.close();
              notesWindow = null;
            }
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Closing existing presentation window');
              presentationWindow.removeAllListeners('closed');
              presentationWindow.close();
              presentationWindow = null;
            }
            currentSlide = null;
          } catch (error) {
            console.error('[API] Error closing existing windows:', error.message);
          }
          
          // Load preferences for monitor selection
          const prefs = loadPreferences();
          const displays = screen.getAllDisplays();
          
          const presentationDisplayId = Number(prefs.presentationDisplayId);
          const notesDisplayId = Number(prefs.notesDisplayId);
          
          const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
          const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];
          
          console.log('[API] Using presentation display:', presentationDisplay.id);
          console.log('[API] Using notes display:', notesDisplay.id);
          
          // Open the presentation using the same logic as the IPC handler
          // Create the presentation window
          presentationWindow = new BrowserWindow({
            x: presentationDisplay.bounds.x,
            y: presentationDisplay.bounds.y,
            width: presentationDisplay.bounds.width,
            height: presentationDisplay.bounds.height,
            fullscreen: true,
            frame: false,
            webPreferences: {
              nodeIntegration: false,
              contextIsolation: true,
              partition: GOOGLE_SESSION_PARTITION
            }
          });
          
          // Set up window open handler for speaker notes popup
          presentationWindow.webContents.setWindowOpenHandler(({ url, frameName, features }) => {
            console.log('[API] Window open intercepted:', url);
            console.log('[API] Frame name:', frameName);
            console.log('[API] Features:', features);
            
            // Use default size from Google Slides (no size/position override)
            const windowOptions = {
              frame: false,
              webPreferences: {
                nodeIntegration: false,
                contextIsolation: true,
                partition: GOOGLE_SESSION_PARTITION
              }
            };
            
            return {
              action: 'allow',
              overrideBrowserWindowOptions: windowOptions
            };
          });
          
          // Listen for notes window creation
          const windowCreatedListener = (event, window) => {
            if (window !== presentationWindow && window !== mainWindow) {
              console.log('[API] Notes window created');
              notesWindow = window;
              
              // Add Escape key handler to notes window
              window.webContents.on('before-input-event', (event, input) => {
                if (input.key === 'Escape' && input.type === 'keyDown') {
                  console.log('[API] Escape pressed in notes window, closing all windows');
                  event.preventDefault();
                  if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
                  if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
                }
              });
              
              // Set speaker notes window to fullscreen when it loads
              window.webContents.once('did-finish-load', () => {
                setSpeakerNotesFullscreen(window);
              });
              
              // Also try when DOM is ready (in case did-finish-load fires too early)
              window.webContents.once('dom-ready', () => {
                setTimeout(() => {
                  setSpeakerNotesFullscreen(window);
                }, 500);
              });
              
              app.removeListener('browser-window-created', windowCreatedListener);
            }
          };
          app.on('browser-window-created', windowCreatedListener);
          
          // Navigation listener (no auto-launch of notes - user must manually start notes)
          const navigationListener = async (event, navUrl) => {
            console.log('[API] Navigated to:', navUrl);
            // Just log navigation, don't auto-launch notes
          };
          
          presentationWindow.webContents.on('did-navigate', navigationListener);
          
          // Listen for page load
          presentationWindow.webContents.once('did-finish-load', async () => {
            console.log('[API] Page finished loading');
            if (!presentationWindow || presentationWindow.isDestroyed()) {
              console.log('[API] Window destroyed before processing');
              return;
            }
            
            await new Promise(resolve => setTimeout(resolve, 200));
            
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Triggering Ctrl+Shift+F5 for presentation mode');
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 50));
              
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
            }
            
            // No auto-launch of speaker notes - user must call /api/open-speaker-notes separately
          });
          
          presentationWindow.on('closed', () => {
            presentationWindow = null;
            currentSlide = null;
          });
          
          // Escape key handler for presentation window
          presentationWindow.webContents.on('before-input-event', (event, input) => {
          if (input.key === 'Escape' && input.type === 'keyDown') {
            event.preventDefault();
            if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
            if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
          }
          });
          
          const presentUrl = toPresentUrl(url);
          console.log('[API] Loading PRESENT URL:', presentUrl);
          lastPresentationUrl = url; // Store original URL (not /present URL) for reload
          currentSlide = 1;
          presentationWindow.loadURL(presentUrl);
          presentationWindow.show();
          // Send response immediately
          if (!res.headersSent) {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Presentation opened (notes not auto-started)' }));
          }
        } catch (error) {
          console.error('[API] Error:', error);
          if (!res.headersSent) {
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: error.message }));
          }
        }
      });
      return;
    }

    // POST /api/open-presentation-with-notes - Open a presentation and automatically launch speaker notes
    if (req.method === 'POST' && req.url === '/api/open-presentation-with-notes') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      
      req.on('end', async () => {
        try {
          const data = JSON.parse(body);
          const { url } = data;
          
          if (!url) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'URL is required' }));
            return;
          }
          
          console.log('[API] Opening presentation with notes:', url);
          
          // Close any existing presentation windows
          try {
            if (notesWindow && !notesWindow.isDestroyed()) {
              console.log('[API] Closing existing notes window');
              notesWindow.removeAllListeners('closed');
              notesWindow.close();
              notesWindow = null;
            }
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Closing existing presentation window');
              presentationWindow.removeAllListeners('closed');
              presentationWindow.close();
              presentationWindow = null;
            }
            currentSlide = null;
          } catch (error) {
            console.error('[API] Error closing existing windows:', error.message);
          }
          
          // Load preferences for monitor selection
          const prefs = loadPreferences();
          const displays = screen.getAllDisplays();
          
          const presentationDisplayId = Number(prefs.presentationDisplayId);
          const notesDisplayId = Number(prefs.notesDisplayId);
          
          const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
          const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];
          
          console.log('[API] Using presentation display:', presentationDisplay.id);
          console.log('[API] Using notes display:', notesDisplay.id);
          
          // Create the presentation window
          presentationWindow = new BrowserWindow({
            x: presentationDisplay.bounds.x,
            y: presentationDisplay.bounds.y,
            width: presentationDisplay.bounds.width,
            height: presentationDisplay.bounds.height,
            fullscreen: true,
            frame: false,
            webPreferences: {
              nodeIntegration: false,
              contextIsolation: true,
              partition: GOOGLE_SESSION_PARTITION
            }
          });
          
          // Set up window open handler for speaker notes popup
          presentationWindow.webContents.setWindowOpenHandler(({ url, frameName, features }) => {
            // Use default size from Google Slides (no size/position override)
            const windowOptions = {
              frame: false,
              webPreferences: {
                nodeIntegration: false,
                contextIsolation: true,
                partition: GOOGLE_SESSION_PARTITION
              }
            };
            
            return {
              action: 'allow',
              overrideBrowserWindowOptions: windowOptions
            };
          });
          
          // Listen for notes window creation
          const windowCreatedListener = (event, window) => {
            if (window !== presentationWindow && window !== mainWindow) {
              console.log('[API] Notes window created');
              notesWindow = window;
              
              // Add Escape key handler to notes window
              window.webContents.on('before-input-event', (event, input) => {
                if (input.key === 'Escape' && input.type === 'keyDown') {
                  console.log('[API] Escape pressed in notes window, closing all windows');
                  event.preventDefault();
                  if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
                  if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
                }
              });
              
              // Set speaker notes window to fullscreen when it loads
              window.webContents.once('did-finish-load', () => {
                setSpeakerNotesFullscreen(window);
              });
              
              // Also try when DOM is ready (in case did-finish-load fires too early)
              window.webContents.once('dom-ready', () => {
                setTimeout(() => {
                  setSpeakerNotesFullscreen(window);
                }, 500);
              });
              
              app.removeListener('browser-window-created', windowCreatedListener);
            }
          };
          app.on('browser-window-created', windowCreatedListener);
          
          // Navigation listener - auto-launch speaker notes when in presentation mode
          let sKeyPressed = false;
          const navigationListener = async (event, navUrl) => {
            console.log('[API] Navigated to:', navUrl);
            
            // Check if we're in presentation mode (URL contains /present/ or /localpresent but not /presentation/)
            const isPresentMode = (navUrl.includes('/present/') || navUrl.includes('/localpresent')) && !navUrl.includes('/presentation/');
            if (isPresentMode && !sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
              sKeyPressed = true;
              console.log('[API] Presentation mode detected, auto-launching speaker notes...');
              
              // Small delay to ensure presentation mode UI is ready
              await new Promise(resolve => setTimeout(resolve, 300));
              
              if (presentationWindow && !presentationWindow.isDestroyed()) {
                try {
                  presentationWindow.focus();
                  await new Promise(resolve => setTimeout(resolve, 50));
                  
                  presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
                  presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
                  presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
                  
                  console.log('[API] "s" key sent for speaker notes');
                } catch (error) {
                  console.error('[API] Error sending "s" key:', error);
                }
                
                presentationWindow.webContents.removeListener('did-navigate', navigationListener);
              }
            }
          };
          
          presentationWindow.webContents.on('did-navigate', navigationListener);
          
          // Listen for page load
          presentationWindow.webContents.once('did-finish-load', async () => {
            console.log('[API] Page finished loading');
            if (!presentationWindow || presentationWindow.isDestroyed()) {
              console.log('[API] Window destroyed before processing');
              return;
            }
            
            await new Promise(resolve => setTimeout(resolve, 200));
            
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Triggering Ctrl+Shift+F5 for presentation mode');
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 50));
              
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
            }
            
            // Fallback: press 's' for speaker notes after delay if navigation didn't trigger it
            setTimeout(async () => {
              if (!sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
                console.log('[API] Fallback: pressing "s" for speaker notes');
                sKeyPressed = true;
                presentationWindow.focus();
                await new Promise(resolve => setTimeout(resolve, 50));
                
                presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
                presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
                presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
                
                if (presentationWindow && !presentationWindow.isDestroyed()) {
                  presentationWindow.webContents.removeListener('did-navigate', navigationListener);
                }
              }
            }, 1000);
          });
          
          presentationWindow.on('closed', () => {
            presentationWindow = null;
            currentSlide = null;
          });
          
          // Escape key handler for presentation window
          presentationWindow.webContents.on('before-input-event', (event, input) => {
          if (input.key === 'Escape' && input.type === 'keyDown') {
            event.preventDefault();
            if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
            if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
          }
          });
          
          const presentUrl = toPresentUrl(url);
          console.log('[API] Loading PRESENT URL:', presentUrl);
          lastPresentationUrl = url; // Store original URL (not /present URL) for reload
          currentSlide = 1;
          presentationWindow.loadURL(presentUrl);
          presentationWindow.show();
          
          // Send response immediately
          if (!res.headersSent) {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Presentation opened with notes' }));
          }
        } catch (error) {
          console.error('[API] Error opening presentation with notes:', error);
          if (!res.headersSent) {
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: error.message }));
          }
        }
      });
      return;
    }
    
    // POST /api/close-presentation - Close current presentation
    if (req.method === 'POST' && req.url === '/api/close-presentation') {
      console.log('[API] Closing presentation');
      
      // Send response first before closing windows
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, message: 'Presentation closed' }));
      
      // Close windows after sending response to avoid errors
      setImmediate(() => {
        try {
          if (notesWindow && !notesWindow.isDestroyed()) {
            notesWindow.removeAllListeners('closed');
            notesWindow.close();
            notesWindow = null;
          }
          if (presentationWindow && !presentationWindow.isDestroyed()) {
            presentationWindow.removeAllListeners('closed');
            presentationWindow.close();
            presentationWindow = null;
          }
          currentSlide = null;
        } catch (error) {
          console.error('[API] Error closing windows:', error.message);
        }
      });
      
      return;
    }
    
    // POST /api/next-slide - Go to next slide
    if (req.method === 'POST' && req.url === '/api/next-slide') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }
      
      try {
        presentationWindow.focus();
        presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Right' });
        presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Right' });
        currentSlide = (typeof currentSlide === 'number' ? currentSlide + 1 : 1);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Next slide' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // POST /api/previous-slide - Go to previous slide
    if (req.method === 'POST' && req.url === '/api/previous-slide') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }
      
      try {
        presentationWindow.focus();
        presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Left' });
        presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Left' });
        currentSlide = (typeof currentSlide === 'number' && currentSlide > 1 ? currentSlide - 1 : 1);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Previous slide' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/go-to-slide - Navigate to a specific slide number
    if (req.method === 'POST' && req.url === '/api/go-to-slide') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }

      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });

      req.on('end', async () => {
        try {
          const data = JSON.parse(body);
          const targetSlide = parseInt(data.slide, 10);

          if (isNaN(targetSlide) || targetSlide < 1) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'Valid slide number (>= 1) is required' }));
            return;
          }

          // Get current slide (from our tracking or default to 1)
          const current = typeof currentSlide === 'number' ? currentSlide : 1;
          const slidesToMove = targetSlide - current;

          if (slidesToMove === 0) {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Already on slide ' + targetSlide }));
            return;
          }

          presentationWindow.focus();
          await new Promise(resolve => setTimeout(resolve, 50));

          // Send arrow key presses to navigate
          const keyCode = slidesToMove > 0 ? 'Right' : 'Left';
          const count = Math.abs(slidesToMove);

          for (let i = 0; i < count; i++) {
            presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: keyCode });
            presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: keyCode });
            // Small delay between key presses to ensure they're processed
            if (i < count - 1) {
              await new Promise(resolve => setTimeout(resolve, 100));
            }
          }

          // Update our tracking
          currentSlide = targetSlide;

          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ 
            success: true, 
            message: `Navigated to slide ${targetSlide}`,
            fromSlide: current,
            toSlide: targetSlide
          }));
        } catch (error) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        }
      });
      return;
    }

    // POST /api/reload-presentation - Close, reopen, and return to current slide
    if (req.method === 'POST' && req.url === '/api/reload-presentation') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }

      if (!lastPresentationUrl) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No previous presentation URL stored' }));
        return;
      }

      // Send response immediately (this will be async)
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, message: 'Reloading presentation...' }));

      // Do the reload asynchronously
      (async () => {
        try {
          // Step 1: Remember current slide number
          const savedSlide = typeof currentSlide === 'number' ? currentSlide : 1;
          
          // Step 2: Remember if speaker notes are open (boolean)
          const notesWereOpen = !!(notesWindow && !notesWindow.isDestroyed());
          
          // Step 3: Remember the URL
          const urlToReload = lastPresentationUrl;
          
          console.log('[API] Reload: Saving state - slide:', savedSlide, 'notes open:', notesWereOpen, 'URL:', urlToReload);
          
          // Step 4: Close the presentation
          if (notesWindow && !notesWindow.isDestroyed()) {
            notesWindow.removeAllListeners('closed');
            notesWindow.close();
            notesWindow = null;
          }
          if (presentationWindow && !presentationWindow.isDestroyed()) {
            presentationWindow.removeAllListeners('closed');
            presentationWindow.close();
            presentationWindow = null;
          }
          currentSlide = null;

          // Wait for windows to close
          await new Promise(resolve => setTimeout(resolve, 200));

          // Step 5: Reopen the presentation (using same logic as /api/open-presentation)
          const prefs = loadPreferences();
          const displays = screen.getAllDisplays();
          const presentationDisplayId = Number(prefs.presentationDisplayId);
          const notesDisplayId = Number(prefs.notesDisplayId);
          const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
          const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];

          presentationWindow = new BrowserWindow({
            x: presentationDisplay.bounds.x,
            y: presentationDisplay.bounds.y,
            width: presentationDisplay.bounds.width,
            height: presentationDisplay.bounds.height,
            fullscreen: true,
            frame: false,
            webPreferences: {
              nodeIntegration: false,
              contextIsolation: true,
              partition: GOOGLE_SESSION_PARTITION
            }
          });

          // Set up window open handler for speaker notes popup
          presentationWindow.webContents.setWindowOpenHandler(({ url, frameName, features }) => {
            // Use default size from Google Slides (no size/position override)
            const windowOptions = {
              frame: false,
              webPreferences: {
                nodeIntegration: false,
                contextIsolation: true,
                partition: GOOGLE_SESSION_PARTITION
              }
            };
            
            return {
              action: 'allow',
              overrideBrowserWindowOptions: windowOptions
            };
          });
          
          // Listen for notes window creation
          const windowCreatedListener = (event, window) => {
            if (window !== presentationWindow && window !== mainWindow) {
              console.log('[API] Reload: Notes window created');
              notesWindow = window;
              
              // Add Escape key handler to notes window
              window.webContents.on('before-input-event', (event, input) => {
                if (input.key === 'Escape' && input.type === 'keyDown') {
                  event.preventDefault();
                  if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
                  if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
                }
              });
              
              // Set speaker notes window to fullscreen when it loads
              window.webContents.once('did-finish-load', () => {
                setSpeakerNotesFullscreen(window);
              });
              
              // Also try when DOM is ready (in case did-finish-load fires too early)
              window.webContents.once('dom-ready', () => {
                setTimeout(() => {
                  setSpeakerNotesFullscreen(window);
                }, 500);
              });
              
              app.removeListener('browser-window-created', windowCreatedListener);
            }
          };
          app.on('browser-window-created', windowCreatedListener);
          
          // Set up window event handlers BEFORE loading URL
          presentationWindow.on('closed', () => {
            presentationWindow = null;
            currentSlide = null;
          });
          
          // Escape key handler for presentation window - set up BEFORE showing window
          presentationWindow.webContents.on('before-input-event', (event, input) => {
            if (input.key === 'Escape' && input.type === 'keyDown') {
              console.log('[API] Reload: Escape pressed, closing presentation and notes windows');
              event.preventDefault();
              if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
              if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
            }
          });

          // Load the presentation
          const presentUrl = toPresentUrl(urlToReload);
          lastPresentationUrl = urlToReload;
          currentSlide = 1; // Will be updated after navigation
          
          presentationWindow.loadURL(presentUrl);
          
          // Show and focus the window immediately to ensure it receives keyboard events
          presentationWindow.show();
          presentationWindow.focus();
          
          // Also focus when window becomes ready
          presentationWindow.once('ready-to-show', () => {
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              presentationWindow.focus();
            }
          });
          
          // Trigger presentation mode when page loads
          presentationWindow.webContents.once('did-finish-load', async () => {
            console.log('[API] Reload: Page finished loading');
            if (!presentationWindow || presentationWindow.isDestroyed()) {
              return;
            }
            
            // Ensure window is focused before sending keyboard events
            presentationWindow.focus();
            await new Promise(resolve => setTimeout(resolve, 200));
            
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Reload: Triggering Ctrl+Shift+F5 for presentation mode');
              // Focus again before sending keyboard event
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 100));
              
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
            }
          });

          // Wait for presentation mode to be ready
          await new Promise(resolve => setTimeout(resolve, 2000));
          
          // Step 6: Navigate to the saved slide number using the go-to-slide endpoint
          if (presentationWindow && !presentationWindow.isDestroyed() && savedSlide > 1) {
            console.log('[API] Reload: Navigating to saved slide:', savedSlide);
            
            // Use the go-to-slide endpoint to jump directly to the slide
            const prefs = loadPreferences();
            const apiPort = prefs.apiPort || DEFAULT_API_PORT;
            
            try {
              const postData = JSON.stringify({ slide: savedSlide });
              
              const options = {
                hostname: '127.0.0.1',
                port: apiPort,
                path: '/api/go-to-slide',
                method: 'POST',
                headers: {
                  'Content-Type': 'application/json',
                  'Content-Length': Buffer.byteLength(postData)
                }
              };
              
              await new Promise((resolve, reject) => {
                const req = http.request(options, (res) => {
                  let data = '';
                  res.on('data', (chunk) => { data += chunk; });
                  res.on('end', () => {
                    try {
                      const result = JSON.parse(data);
                      if (result.success) {
                        console.log('[API] Reload: Successfully navigated to slide', savedSlide);
                      } else {
                        console.error('[API] Reload: Failed to navigate to slide:', result.error);
                      }
                      resolve();
                    } catch (err) {
                      console.error('[API] Reload: Error parsing response:', err);
                      resolve();
                    }
                  });
                });
                
                req.on('error', (err) => {
                  console.error('[API] Reload: Error calling go-to-slide endpoint:', err);
                  resolve(); // Don't reject, just log the error
                });
                
                req.write(postData);
                req.end();
              });
            } catch (error) {
              console.error('[API] Reload: Error calling go-to-slide endpoint:', error);
            }
          } else if (savedSlide === 1) {
            // If we're already on slide 1, just update the tracking
            currentSlide = 1;
            console.log('[API] Reload: Already on slide 1');
          }
          
          // Step 7: Reopen speaker notes if they were previously open
          if (notesWereOpen && presentationWindow && !presentationWindow.isDestroyed()) {
            console.log('[API] Reload: Speaker notes were previously open, launching them now');
            await new Promise(resolve => setTimeout(resolve, 1000)); // Wait longer for presentation mode to be ready
            
            // Ensure window has focus before sending keyboard events
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 200));
              
              // Send the 's' key to open speaker notes
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
              presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
              
              console.log('[API] Reload: Speaker notes launch command sent');
            }
          } else {
            console.log('[API] Reload: Speaker notes were not previously open, skipping');
          }
          
          // Final focus to ensure window is ready for keyboard input
          if (presentationWindow && !presentationWindow.isDestroyed()) {
            setTimeout(() => {
              presentationWindow.focus();
            }, 500);
          }
          
          console.log('[API] Reload: Complete');

        } catch (error) {
          console.error('[API] Error during reload:', error);
        }
      })();

      return;
    }
    
    // POST /api/toggle-video - Toggle video playback (k key)
    if (req.method === 'POST' && req.url === '/api/toggle-video') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }
      
      try {
        presentationWindow.focus();
        presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'K' });
        presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 'k' });
        presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'K' });
        
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Video toggled' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/open-speaker-notes - Open/start speaker notes (s key)
    if (req.method === 'POST' && req.url === '/api/open-speaker-notes') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }

      try {
        console.log('[API] Opening speaker notes');
        presentationWindow.focus();
        await new Promise(resolve => setTimeout(resolve, 50));
        presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
        presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
        presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Speaker notes opened' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/close-speaker-notes - Close the speaker notes window
    if (req.method === 'POST' && req.url === '/api/close-speaker-notes') {
      if (!notesWindow || notesWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }

      try {
        notesWindow.close();
        notesWindow = null;
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Speaker notes closed' }));
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/scroll-notes-down - Scroll speaker notes down (JS only, no keyboard)
    if (req.method === 'POST' && req.url === '/api/scroll-notes-down') {
      if (!notesWindow || notesWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }
      try {
        notesWindow.webContents.executeJavaScript(`
          (function() {
            // Find scrollable elements - try common patterns in Google Slides presenter view
            var scrollable = null;
            var allElements = document.querySelectorAll('*');
            for (var i = 0; i < allElements.length; i++) {
              var el = allElements[i];
              var style = window.getComputedStyle(el);
              if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && 
                  el.scrollHeight > el.clientHeight) {
                scrollable = el;
                break;
              }
            }
            // Fallback: try document body or documentElement if they're scrollable
            if (!scrollable) {
              if (document.body && document.body.scrollHeight > document.body.clientHeight) {
                scrollable = document.body;
              } else if (document.documentElement && document.documentElement.scrollHeight > document.documentElement.clientHeight) {
                scrollable = document.documentElement;
              }
            }
            if (scrollable) {
              scrollable.scrollBy(0, 150);
              return { success: true, scrolled: true };
            }
            return { success: false, error: 'No scrollable element found' };
          })()
        `).then(result => {
          if (result.success && result.scrolled) {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Notes scrolled down' }));
          } else {
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: result.error || 'Could not scroll notes' }));
          }
        }).catch(error => {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        });
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }

    // POST /api/scroll-notes-up - Scroll speaker notes up (JS only, no keyboard)
    if (req.method === 'POST' && req.url === '/api/scroll-notes-up') {
      if (!notesWindow || notesWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }
      try {
        notesWindow.webContents.executeJavaScript(`
          (function() {
            // Find scrollable elements - try common patterns in Google Slides presenter view
            var scrollable = null;
            var allElements = document.querySelectorAll('*');
            for (var i = 0; i < allElements.length; i++) {
              var el = allElements[i];
              var style = window.getComputedStyle(el);
              if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && 
                  el.scrollHeight > el.clientHeight) {
                scrollable = el;
                break;
              }
            }
            // Fallback: try document body or documentElement if they're scrollable
            if (!scrollable) {
              if (document.body && document.body.scrollHeight > document.body.clientHeight) {
                scrollable = document.body;
              } else if (document.documentElement && document.documentElement.scrollHeight > document.documentElement.clientHeight) {
                scrollable = document.documentElement;
              }
            }
            if (scrollable) {
              scrollable.scrollBy(0, -150);
              return { success: true, scrolled: true };
            }
            return { success: false, error: 'No scrollable element found' };
          })()
        `).then(result => {
          if (result.success && result.scrolled) {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Notes scrolled up' }));
          } else {
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: result.error || 'Could not scroll notes' }));
          }
        }).catch(error => {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        });
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // POST /api/zoom-in-notes - Zoom in on speaker notes
    if (req.method === 'POST' && req.url === '/api/zoom-in-notes') {
      console.log('[API] Zoom in on speaker notes requested');
      
      if (!notesWindow || notesWindow.isDestroyed()) {
        console.log('[API] No speaker notes window is open');
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }
      
      try {
        notesWindow.webContents.executeJavaScript(`
          (function() {
            const zoomInButton = document.querySelector('[title="Zoom in"]');
            if (zoomInButton) {
              // Dispatch real mouse events
              const mousedownEvent = new MouseEvent('mousedown', {
                bubbles: true,
                cancelable: true,
                view: window,
                button: 0
              });
              const mouseupEvent = new MouseEvent('mouseup', {
                bubbles: true,
                cancelable: true,
                view: window,
                button: 0
              });
              const clickEvent = new MouseEvent('click', {
                bubbles: true,
                cancelable: true,
                view: window,
                button: 0
              });
              
              zoomInButton.dispatchEvent(mousedownEvent);
              zoomInButton.dispatchEvent(mouseupEvent);
              zoomInButton.dispatchEvent(clickEvent);
              
              return { success: true };
            }
            return { success: false, error: 'Button not found' };
          })()
        `).then(result => {
          if (result.success) {
            console.log('[API] ✓ Dispatched mouse events to zoom in button');
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Zoomed in on notes' }));
          } else {
            console.log('[API] ✗ Zoom in button not found');
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: result.error }));
          }
        }).catch(error => {
          console.error('[API] Error executing zoom in script:', error.message);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        });
      } catch (error) {
        console.error('[API] Error zooming in on notes:', error.message);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // POST /api/zoom-out-notes - Zoom out on speaker notes
    if (req.method === 'POST' && req.url === '/api/zoom-out-notes') {
      console.log('[API] Zoom out on speaker notes requested');
      
      if (!notesWindow || notesWindow.isDestroyed()) {
        console.log('[API] No speaker notes window is open');
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No speaker notes window is open' }));
        return;
      }
      
      try {
        notesWindow.webContents.executeJavaScript(`
          (function() {
            const zoomOutButton = document.querySelector('[title="Zoom out"]');
            if (zoomOutButton) {
              // Dispatch real mouse events
              const mousedownEvent = new MouseEvent('mousedown', {
                bubbles: true,
                cancelable: true,
                view: window,
                button: 0
              });
              const mouseupEvent = new MouseEvent('mouseup', {
                bubbles: true,
                cancelable: true,
                view: window,
                button: 0
              });
              const clickEvent = new MouseEvent('click', {
                bubbles: true,
                cancelable: true,
                view: window,
                button: 0
              });
              
              zoomOutButton.dispatchEvent(mousedownEvent);
              zoomOutButton.dispatchEvent(mouseupEvent);
              zoomOutButton.dispatchEvent(clickEvent);
              
              return { success: true };
            }
            return { success: false, error: 'Button not found' };
          })()
        `).then(result => {
          if (result.success) {
            console.log('[API] ✓ Dispatched mouse events to zoom out button');
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Zoomed out on notes' }));
          } else {
            console.log('[API] ✗ Zoom out button not found');
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: result.error }));
          }
        }).catch(error => {
          console.error('[API] Error executing zoom out script:', error.message);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        });
      } catch (error) {
        console.error('[API] Error zooming out on notes:', error.message);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
      return;
    }
    
    // GET /api/get-speaker-notes - Get current speaker notes content
    if (req.method === 'GET' && req.url === '/api/get-speaker-notes') {
      if (!notesWindow || notesWindow.isDestroyed()) {
        res.writeHead(200, { 
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        });
        res.end(JSON.stringify({ success: false, notes: '', error: 'No speaker notes window is open' }));
        return;
      }

      (async () => {
        try {
          const notesContent = await notesWindow.webContents.executeJavaScript(`
            (function(){
              // First, try to find the specific div that contains only the notes text
              var notesTextDiv = document.querySelector('div.punch-viewer-speakernotes-text-body-scrollable');
              
              if (notesTextDiv) {
                // This is the exact element we want - just get its text
                var notesText = notesTextDiv.innerText || notesTextDiv.textContent || '';
                notesText = notesText.trim();
                
                if (notesText.length > 0) {
                  return notesText;
                }
              }
              
              // Fallback: Look for the table if the specific div isn't found
              var notesTable = document.querySelector('table.punch-viewer-speakernotes, table[id*="speakernotes"], table[class*="speakernotes"]');
              
              if (!notesTable) {
                // Try alternative selectors
                notesTable = document.querySelector('[class*="punch-viewer-speakernotes"]');
              }
              
              if (!notesTable) {
                // Try finding by ID
                notesTable = document.getElementById('punch-viewer-speakernotes');
              }
              
              if (notesTable) {
                // Try to find the text body div within the table
                var textBodyDiv = notesTable.querySelector('div.punch-viewer-speakernotes-text-body-scrollable');
                if (textBodyDiv) {
                  var notesText = textBodyDiv.innerText || textBodyDiv.textContent || '';
                  notesText = notesText.trim();
                  
                  if (notesText.length > 0) {
                    return notesText;
                  }
                }
                
                // Last resort: Clone the table and clean it up
                var tableClone = notesTable.cloneNode(true);
                
                // Remove UI elements that shouldn't be in notes
                var uiElements = tableClone.querySelectorAll('button, [role="button"], [class*="button"], [class*="control"], [class*="timer"], [class*="thumbnail"], [aria-label*="Pause"], [aria-label*="Reset"], [aria-label*="Next"], [aria-label*="Previous"]');
                for (var i = 0; i < uiElements.length; i++) {
                  uiElements[i].remove();
                }
                
                // Remove cells that are clearly UI (short text, button labels, timers)
                var allCells = tableClone.querySelectorAll('td, th');
                for (var j = 0; j < allCells.length; j++) {
                  var cell = allCells[j];
                  var cellText = (cell.innerText || cell.textContent || '').trim();
                  
                  if (cellText.length <= 2) {
                    cell.remove();
                    continue;
                  }
                  if (/^(Pause|Reset|Next|Previous|Zoom|Timer|Slide)$/i.test(cellText)) {
                    cell.remove();
                    continue;
                  }
                  if (/^\\d{1,2}:\\d{2}(?::\\d{2})?$/.test(cellText)) {
                    cell.remove();
                    continue;
                  }
                  if (/^\\d+$/.test(cellText) && cellText.length <= 3) {
                    cell.remove();
                    continue;
                  }
                }
                
                var notesText = tableClone.innerText || tableClone.textContent || '';
                notesText = notesText.trim();
                
                // Remove common UI patterns that might leak through
                // Only remove these in specific UI contexts, not when they're part of actual notes
                notesText = notesText.replace(/Slide \\d+ of \\d+/gi, '');
                notesText = notesText.replace(/^Slide \\d+$/gmi, ''); // Remove standalone "Slide X" lines
                notesText = notesText.replace(/\\d{1,2}:\\d{2}(?::\\d{2})?/g, '');
                // Only remove specific UI button labels (not common words that might be in notes)
                notesText = notesText.replace(/\\b(Previous Slide|Next Slide)\\b/gi, '');
                notesText = notesText.replace(/\\b(Pause|Reset)\\b/gi, '');
                notesText = notesText.replace(/\\b(Zoom In|Zoom Out)\\b/gi, '');
                notesText = notesText.replace(/AUDIENCE TOOLS|SPEAKER NOTES/gi, '');
                notesText = notesText.replace(/Q&A|Questions|Answers/gi, '');
                
                // Split into lines and filter out slide numbers and blank lines at the start
                var lines = notesText.split('\\n');
                var startIndex = 0;
                // Skip leading blank lines and slide number patterns
                for (var i = 0; i < lines.length; i++) {
                  var line = lines[i].trim();
                  // Skip blank lines
                  if (line === '') {
                    continue;
                  }
                  // Skip lines that are just slide numbers (e.g., "1", "Slide 1", "1 of 10")
                  if (/^\\d+$/.test(line) || /^Slide \\d+/i.test(line) || /^\\d+ of \\d+$/i.test(line)) {
                    continue;
                  }
                  // Found first real content line
                  startIndex = i;
                  break;
                }
                // Rejoin from the first real content line
                notesText = lines.slice(startIndex).join('\\n');
                
                // Clean up multiple newlines and whitespace
                notesText = notesText.replace(/\\n{3,}/g, '\\n\\n');
                notesText = notesText.replace(/[ \\t]{2,}/g, ' ');
                notesText = notesText.trim();
                
                if (notesText.length > 0) {
                  return notesText;
                }
              }
              
              // Fallback: try to find any element with that class/id pattern
              var notesElement = document.querySelector('[class*="punch-viewer-speakernotes"], [id*="speakernotes"]');
              if (notesElement) {
                // First, try to find the specific text body div within this element
                var textBodyDiv = notesElement.querySelector('div.punch-viewer-speakernotes-text-body-scrollable');
                if (textBodyDiv) {
                  var notesText = textBodyDiv.innerText || textBodyDiv.textContent || '';
                  notesText = notesText.trim();
                  
                  if (notesText.length > 0) {
                    return notesText;
                  }
                }
                
                // If we didn't find the specific div, fall back to cleaning the element
                var elementClone = notesElement.cloneNode(true);
                
                // Remove UI elements
                var uiElements = elementClone.querySelectorAll('button, [role="button"], [class*="button"], [class*="control"], [class*="timer"], [class*="thumbnail"], [aria-label*="Pause"], [aria-label*="Reset"], [aria-label*="Next"], [aria-label*="Previous"]');
                for (var i = 0; i < uiElements.length; i++) {
                  uiElements[i].remove();
                }
                
                // If it's a table, clean up cells like we did above
                if (elementClone.tagName === 'TABLE') {
                  var allCells = elementClone.querySelectorAll('td, th');
                  for (var j = 0; j < allCells.length; j++) {
                    var cell = allCells[j];
                    var cellText = (cell.innerText || cell.textContent || '').trim();
                    
                    if (cellText.length <= 2) {
                      cell.remove();
                      continue;
                    }
                    if (/^(Pause|Reset|Next|Previous|Zoom|Timer|Slide)$/i.test(cellText)) {
                      cell.remove();
                      continue;
                    }
                    if (/^\\d{1,2}:\\d{2}(?::\\d{2})?$/.test(cellText)) {
                      cell.remove();
                      continue;
                    }
                    if (/^\\d+$/.test(cellText) && cellText.length <= 3) {
                      cell.remove();
                      continue;
                    }
                  }
                }
                
                var notesText = elementClone.innerText || elementClone.textContent || '';
                notesText = notesText.trim();
                
                // Clean up remaining patterns
                notesText = notesText.replace(/Slide \\d+ of \\d+/gi, '');
                notesText = notesText.replace(/^Slide \\d+$/gmi, '');
                notesText = notesText.replace(/AUDIENCE TOOLS|SPEAKER NOTES/gi, '');
                notesText = notesText.replace(/Q&A|Questions|Answers/gi, '');
                
                // Split into lines and filter out slide numbers and blank lines at the start
                var lines = notesText.split('\\n');
                var startIndex = 0;
                for (var k = 0; k < lines.length; k++) {
                  var line = lines[k].trim();
                  if (line === '') continue;
                  if (/^\\d+$/.test(line) || /^Slide \\d+/i.test(line) || /^\\d+ of \\d+$/i.test(line)) {
                    continue;
                  }
                  startIndex = k;
                  break;
                }
                notesText = lines.slice(startIndex).join('\\n');
                notesText = notesText.replace(/\\n{3,}/g, '\\n\\n').trim();
                
                if (notesText.length > 0) {
                  return notesText;
                }
              }
              
              return 'No notes available for this slide.';
            })()
          `);
          
          res.writeHead(200, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ success: true, notes: notesContent || 'No notes available for this slide.' }));
        } catch (error) {
          console.error('[API] Error getting speaker notes:', error);
          res.writeHead(200, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ success: false, notes: '', error: error.message }));
        }
      })();
      return;
    }
    
    // GET /api/get-stagetimer-status - Get live timer data from stagetimer.io
    if (req.method === 'GET' && req.url === '/api/get-stagetimer-status') {
      const prefs = loadPreferences();
      const roomId = prefs.stagetimerRoomId;
      const apiKey = prefs.stagetimerApiKey;
      
      if (!roomId || !apiKey) {
        res.writeHead(200, { 
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        });
        res.end(JSON.stringify({ 
          success: false, 
          error: 'Stagetimer not configured. Please set Room ID and API Key in Settings.',
          configured: false
        }));
        return;
      }
      
      // Call stagetimer.io API for status
      const statusUrl = `https://api.stagetimer.io/v1/get_status?room_id=${encodeURIComponent(roomId)}&api_key=${encodeURIComponent(apiKey)}`;
      const messagesUrl = `https://api.stagetimer.io/v1/get_all_messages?room_id=${encodeURIComponent(roomId)}&api_key=${encodeURIComponent(apiKey)}`;
      
      // Fetch status and messages in parallel, then fetch timer using timer_id
      let statusData = null;
      let messagesData = null;
      let timerData = null;
      let completed = 0;
      let timerCompleted = false;
      let timerTimeout = null;
      const totalRequests = 2; // Status and messages first
      const TIMER_FETCH_TIMEOUT = 5000; // 5 second timeout for timer fetch
      
      function sendResponse() {
        // Wait for status and messages
        if (completed < totalRequests) return;
        // If we're waiting for timer, give it a chance, but don't wait forever
        if (timerData === null && !timerCompleted) {
          // Set a timeout to send response even if timer fetch hangs
          if (!timerTimeout) {
            timerTimeout = setTimeout(() => {
              console.warn('[API] Timer fetch timeout, sending response without timer data');
              timerCompleted = true;
              timerData = { ok: false, data: {}, error: 'Timeout' };
              sendResponse();
            }, TIMER_FETCH_TIMEOUT);
          }
          return;
        }
        
        // Clear timeout if we got here normally
        if (timerTimeout) {
          clearTimeout(timerTimeout);
          timerTimeout = null;
        }
        
        try {
          if (statusData && statusData.ok && statusData.data) {
            const status = statusData.data;
            const now = status.server_time || Date.now();
            
            // Calculate remaining/elapsed time
            let remainingMs = 0;
            let elapsedMs = 0;
            let displayTime = '0:00';
            let isRunning = status.running || false;
            
            if (status.finish && status.start) {
              const duration = status.finish - status.start;
              
              if (isRunning) {
                remainingMs = status.finish - now; // Allow negative values
                elapsedMs = now - status.start;
              } else if (status.pause) {
                // Timer is paused
                elapsedMs = status.pause - status.start;
                remainingMs = duration - elapsedMs;
              } else {
                // Timer not started
                remainingMs = duration;
                elapsedMs = 0;
              }
              
              // Format time as MM:SS or HH:MM:SS (allow negative)
              const totalSeconds = Math.floor(remainingMs / 1000);
              const isNegative = totalSeconds < 0;
              const absSeconds = Math.abs(totalSeconds);
              const hours = Math.floor(absSeconds / 3600);
              const minutes = Math.floor((absSeconds % 3600) / 60);
              const seconds = absSeconds % 60;
              
              const sign = isNegative ? '-' : '';
              const minStr = String(minutes).padStart(2, '0');
              const secStr = String(seconds).padStart(2, '0');
              
              if (hours > 0) {
                displayTime = sign + hours + ':' + minStr + ':' + secStr;
              } else {
                displayTime = sign + minutes + ':' + secStr;
              }
            }
            
            // Process messages
            let activeMessages = [];
            if (messagesData && messagesData.ok && messagesData.data) {
              // Check if messages is an array directly or nested in data.messages
              let messages = [];
              if (Array.isArray(messagesData.data)) {
                messages = messagesData.data;
              } else if (messagesData.data.messages && Array.isArray(messagesData.data.messages)) {
                messages = messagesData.data.messages;
              }
              
              console.log('[API] Processing messages, found:', messages.length, 'total messages');
              activeMessages = messages
                .filter(msg => msg && msg.showing === true)
                .map(msg => ({
                  text: msg.text || '',
                  color: msg.color || 'white',
                  bold: msg.bold || false,
                  uppercase: msg.uppercase || false
                }));
              console.log('[API] Active messages:', activeMessages.length);
            } else {
              console.log('[API] No messages data or not ok:', messagesData);
            }
            
            // Get timer name and speaker from timer data
            let timerName = '';
            let speakerName = '';
            console.log('[API] Processing timer data, timerData:', timerData ? 'exists' : 'null');
            console.log('[API] timerData.ok:', timerData?.ok);
            console.log('[API] timerData.data:', timerData?.data ? JSON.stringify(timerData.data) : 'null');
            if (timerData && timerData.ok && timerData.data) {
              timerName = timerData.data.name || '';
              speakerName = timerData.data.speaker || '';
              console.log('[API] Extracted timerName:', timerName, 'speakerName:', speakerName);
            } else {
              console.log('[API] Timer data not available or invalid. timerData:', timerData);
            }
            
            res.writeHead(200, { 
              'Content-Type': 'application/json',
              'Access-Control-Allow-Origin': '*'
            });
            res.end(JSON.stringify({
              success: true,
              configured: true,
              running: isRunning,
              displayTime: displayTime,
              remainingMs: remainingMs,
              elapsedMs: elapsedMs,
              timerId: status.timer_id || null,
              start: status.start,
              finish: status.finish,
              pause: status.pause,
              serverTime: status.server_time,
              messages: activeMessages,
              timerName: timerName,
              speaker: speakerName,
              // Debug info
              _debug: {
                timerDataOk: timerData?.ok || false,
                timerDataExists: !!timerData,
                timerDataMessage: timerData?.message || null,
                rawTimerName: timerData?.data?.name || null,
                rawSpeaker: timerData?.data?.speaker || null
              }
            }));
          } else {
            // Status fetch failed or returned invalid data
            // Still try to send response with whatever we have
            res.writeHead(200, { 
              'Content-Type': 'application/json',
              'Access-Control-Allow-Origin': '*'
            });
            const errorMsg = statusData 
              ? (statusData.message || statusData.error || 'Failed to get timer status')
              : 'Failed to get timer status';
            console.error('[API] Status fetch failed:', errorMsg, 'statusData:', statusData);
            res.end(JSON.stringify({ 
              success: false, 
              error: errorMsg,
              configured: true,
              _debug: {
                statusDataExists: !!statusData,
                statusDataOk: statusData?.ok,
                statusDataMessage: statusData?.message,
                timerDataExists: !!timerData,
                timerDataOk: timerData?.ok
              }
            }));
          }
        } catch (error) {
          console.error('[API] Error processing stagetimer response:', error);
          res.writeHead(200, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ 
            success: false, 
            error: 'Failed to process response: ' + error.message,
            configured: true
          }));
        }
      }
      
      // Fetch status
      const statusReq = https.get(statusUrl, (apiRes) => {
        let data = '';
        
        // Check HTTP status code
        if (apiRes.statusCode !== 200) {
          console.error('[API] Stagetimer status HTTP error:', apiRes.statusCode, apiRes.statusMessage);
          apiRes.on('data', () => {}); // Drain response
          apiRes.on('end', () => {
            statusData = { ok: false, message: `HTTP ${apiRes.statusCode}: ${apiRes.statusMessage}` };
            completed++;
            sendResponse();
          });
          return;
        }
        
        apiRes.on('data', (chunk) => {
          data += chunk;
        });
        
        apiRes.on('end', () => {
          try {
            statusData = JSON.parse(data);
            if (!statusData || !statusData.ok) {
              console.error('[API] Stagetimer status API error:', statusData?.message || 'Unknown error');
            }
            completed++;
            
            // After getting status, fetch timer using timer_id if available
            if (statusData && statusData.ok && statusData.data && statusData.data.timer_id) {
              const timerId = statusData.data.timer_id;
              const timerUrl = `https://api.stagetimer.io/v1/get_timer?room_id=${encodeURIComponent(roomId)}&api_key=${encodeURIComponent(apiKey)}&timer_id=${encodeURIComponent(timerId)}`;
              
              console.log('[API] Fetching timer with timer_id:', timerId);
              const timerReq = https.get(timerUrl, (timerRes) => {
                let timerDataStr = '';
                
                if (timerRes.statusCode !== 200) {
                  console.error('[API] Stagetimer timer HTTP error:', timerRes.statusCode, timerRes.statusMessage);
                  timerRes.on('data', () => {}); // Drain response
                  timerRes.on('end', () => {
                    if (!timerCompleted) {
                      timerData = { ok: false, message: `HTTP ${timerRes.statusCode}: ${timerRes.statusMessage}`, data: {} };
                      timerCompleted = true;
                      if (timerTimeout) clearTimeout(timerTimeout);
                      sendResponse();
                    }
                  });
                  return;
                }
                
                timerRes.on('data', (chunk) => {
                  timerDataStr += chunk;
                });
                
                timerRes.on('end', () => {
                  if (timerCompleted) return; // Already handled by timeout
                  try {
                    timerData = JSON.parse(timerDataStr);
                    console.log('[API] Stagetimer timer response:', JSON.stringify(timerData, null, 2));
                    console.log('[API] Timer data.name:', timerData?.data?.name);
                    console.log('[API] Timer data.speaker:', timerData?.data?.speaker);
                    timerCompleted = true;
                    if (timerTimeout) clearTimeout(timerTimeout);
                    sendResponse();
                  } catch (error) {
                    console.error('[API] Error parsing stagetimer timer response:', error);
                    console.error('[API] Raw timer response:', timerDataStr);
                    timerData = { ok: false, data: {}, error: error.message };
                    timerCompleted = true;
                    if (timerTimeout) clearTimeout(timerTimeout);
                    sendResponse();
                  }
                });
              });
              
              timerReq.on('error', (error) => {
                if (timerCompleted) return; // Already handled by timeout
                console.error('[API] Error calling stagetimer.io timer:', error);
                timerData = { ok: false, data: {}, error: error.message };
                timerCompleted = true;
                if (timerTimeout) clearTimeout(timerTimeout);
                sendResponse();
              });
              
              // Set request timeout
              timerReq.setTimeout(TIMER_FETCH_TIMEOUT, () => {
                if (!timerCompleted) {
                  console.warn('[API] Timer request timeout');
                  timerReq.destroy();
                  timerData = { ok: false, data: {}, error: 'Request timeout' };
                  timerCompleted = true;
                  if (timerTimeout) clearTimeout(timerTimeout);
                  sendResponse();
                }
              });
            } else {
              // No timer_id available, try without it (gets currently highlighted timer)
              const timerUrl = `https://api.stagetimer.io/v1/get_timer?room_id=${encodeURIComponent(roomId)}&api_key=${encodeURIComponent(apiKey)}`;
              console.log('[API] No timer_id, fetching currently highlighted timer');
              const timerReq = https.get(timerUrl, (timerRes) => {
                let timerDataStr = '';
                
                if (timerRes.statusCode !== 200) {
                  console.error('[API] Stagetimer timer HTTP error:', timerRes.statusCode, timerRes.statusMessage);
                  timerRes.on('data', () => {}); // Drain response
                  timerRes.on('end', () => {
                    if (!timerCompleted) {
                      timerData = { ok: false, message: `HTTP ${timerRes.statusCode}: ${timerRes.statusMessage}`, data: {} };
                      timerCompleted = true;
                      if (timerTimeout) clearTimeout(timerTimeout);
                      sendResponse();
                    }
                  });
                  return;
                }
                
                timerRes.on('data', (chunk) => {
                  timerDataStr += chunk;
                });
                
                timerRes.on('end', () => {
                  if (timerCompleted) return; // Already handled by timeout
                  try {
                    timerData = JSON.parse(timerDataStr);
                    console.log('[API] Stagetimer timer response:', JSON.stringify(timerData, null, 2));
                    console.log('[API] Timer data.name:', timerData?.data?.name);
                    console.log('[API] Timer data.speaker:', timerData?.data?.speaker);
                    timerCompleted = true;
                    if (timerTimeout) clearTimeout(timerTimeout);
                    sendResponse();
                  } catch (error) {
                    console.error('[API] Error parsing stagetimer timer response:', error);
                    console.error('[API] Raw timer response:', timerDataStr);
                    timerData = { ok: false, data: {}, error: error.message };
                    timerCompleted = true;
                    if (timerTimeout) clearTimeout(timerTimeout);
                    sendResponse();
                  }
                });
              });
              
              timerReq.on('error', (error) => {
                if (timerCompleted) return; // Already handled by timeout
                console.error('[API] Error calling stagetimer.io timer:', error);
                timerData = { ok: false, data: {}, error: error.message };
                timerCompleted = true;
                if (timerTimeout) clearTimeout(timerTimeout);
                sendResponse();
              });
              
              // Set request timeout
              timerReq.setTimeout(TIMER_FETCH_TIMEOUT, () => {
                if (!timerCompleted) {
                  console.warn('[API] Timer request timeout');
                  timerReq.destroy();
                  timerData = { ok: false, data: {}, error: 'Request timeout' };
                  timerCompleted = true;
                  if (timerTimeout) clearTimeout(timerTimeout);
                  sendResponse();
                }
              });
            }
            
            sendResponse();
          } catch (error) {
            console.error('[API] Error parsing stagetimer status response:', error);
            statusData = { ok: false, message: 'Failed to parse status response' };
            completed++;
            sendResponse();
          }
        });
      });
      
      statusReq.on('error', (error) => {
        console.error('[API] Error calling stagetimer.io status:', error);
        statusData = { ok: false, message: 'Failed to connect: ' + error.message };
        completed++;
        sendResponse();
      });
      
      // Set request timeout for status
      statusReq.setTimeout(10000, () => {
        console.warn('[API] Status request timeout');
        statusReq.destroy();
        if (!statusData) {
          statusData = { ok: false, message: 'Request timeout' };
          completed++;
          sendResponse();
        }
      });
      
      // Fetch messages
      https.get(messagesUrl, (apiRes) => {
        let data = '';
        
        apiRes.on('data', (chunk) => {
          data += chunk;
        });
        
        apiRes.on('end', () => {
          try {
            messagesData = JSON.parse(data);
            console.log('[API] Stagetimer messages response:', JSON.stringify(messagesData, null, 2));
            completed++;
            sendResponse();
          } catch (error) {
            console.error('[API] Error parsing stagetimer messages response:', error);
            messagesData = { ok: false, data: { messages: [] } };
            completed++;
            sendResponse();
          }
        });
      }).on('error', (error) => {
        console.error('[API] Error calling stagetimer.io messages:', error);
        messagesData = { ok: false, data: { messages: [] } };
        completed++;
        sendResponse();
      });
      
      return;
    }
    
    // GET /api/presets - Get all preset presentations
    if (req.method === 'GET' && req.url === '/api/presets') {
      console.log('[API] GET /api/presets - Loading presets');
      const prefs = loadPreferences();
      console.log('[API] Returning presets:', {
        presentation1: prefs.presentation1 || '',
        presentation2: prefs.presentation2 || '',
        presentation3: prefs.presentation3 || ''
      });
      res.writeHead(200, { 
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
      });
      res.end(JSON.stringify({
        presentation1: prefs.presentation1 || '',
        presentation2: prefs.presentation2 || '',
        presentation3: prefs.presentation3 || ''
      }));
      return;
    }

    // GET /api/debug/preferences - Debug endpoint for preferences file
    if (req.method === 'GET' && req.url === '/api/debug/preferences') {
      try {
        const prefsPath = getPreferencesPath();
        const prefsDir = path.dirname(prefsPath);
        const exists = fs.existsSync(prefsPath);
        const dirExists = fs.existsSync(prefsDir);
        
        let stats = null;
        let content = null;
        let dirWritable = false;
        
        if (exists) {
          stats = fs.statSync(prefsPath);
          try {
            content = fs.readFileSync(prefsPath, 'utf8');
          } catch (e) {
            content = `Error reading file: ${e.message}`;
          }
        }
        
        try {
          fs.accessSync(prefsDir, fs.constants.W_OK);
          dirWritable = true;
        } catch (e) {
          dirWritable = false;
        }
        
        res.writeHead(200, { 
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type'
        });
        res.end(JSON.stringify({
          path: prefsPath,
          directory: prefsDir,
          fileExists: exists,
          directoryExists: dirExists,
          directoryWritable: dirWritable,
          fileSize: stats ? stats.size : null,
          fileModified: stats ? stats.mtime : null,
          fileContent: content,
          preferences: loadPreferences(),
          platform: process.platform,
          userData: app.getPath('userData')
        }));
      } catch (error) {
        res.writeHead(500, { 
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type'
        });
        res.end(JSON.stringify({ error: error.message, stack: error.stack }));
      }
      return;
    }

    // POST /api/stagetimer-settings - Save stagetimer configuration
    if (req.method === 'POST' && req.url === '/api/stagetimer-settings') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      req.on('end', () => {
        try {
          const data = JSON.parse(body);
          const prefs = loadPreferences();
          
          if (data.roomId !== undefined) {
            prefs.stagetimerRoomId = data.roomId || null;
          }
          if (data.apiKey !== undefined) {
            prefs.stagetimerApiKey = data.apiKey || null;
          }
          if (data.enabled !== undefined) {
            prefs.stagetimerEnabled = data.enabled !== false;
          }
          if (data.visible !== undefined) {
            prefs.stagetimerVisible = data.visible !== false;
          } else {
            // Default to true if not set
            prefs.stagetimerVisible = true;
          }
          
          savePreferences(prefs);
          
          res.writeHead(200, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ success: true, message: 'Stagetimer settings saved' }));
        } catch (error) {
          console.error('[API] Error saving stagetimer settings:', error);
          res.writeHead(500, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ success: false, error: error.message }));
        }
      });
      return;
    }
    
    // GET /api/stagetimer-settings - Get stagetimer configuration
    if (req.method === 'GET' && req.url === '/api/stagetimer-settings') {
      const prefs = loadPreferences();
      res.writeHead(200, { 
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      });
      res.end(JSON.stringify({
        roomId: prefs.stagetimerRoomId || '',
        apiKey: prefs.stagetimerApiKey || '',
        enabled: prefs.stagetimerEnabled !== false,
        visible: prefs.stagetimerVisible !== false && prefs.stagetimerVisible !== undefined ? prefs.stagetimerVisible : true
      }));
      return;
    }
    
    // POST /api/presets - Set preset presentations
    if (req.method === 'POST' && req.url === '/api/presets') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      
      req.on('end', () => {
        try {
          console.log('[API] POST /api/presets - Received body:', body);
          const data = JSON.parse(body);
          console.log('[API] Parsed data:', data);
          
          const prefs = loadPreferences();
          console.log('[API] Current preferences before update:', JSON.stringify(prefs));
          
          // Update presets
          if (data.presentation1 !== undefined) {
            prefs.presentation1 = data.presentation1;
            console.log('[API] Updated presentation1:', data.presentation1);
          }
          if (data.presentation2 !== undefined) {
            prefs.presentation2 = data.presentation2;
            console.log('[API] Updated presentation2:', data.presentation2);
          }
          if (data.presentation3 !== undefined) {
            prefs.presentation3 = data.presentation3;
            console.log('[API] Updated presentation3:', data.presentation3);
          }
          
          console.log('[API] Preferences after update:', JSON.stringify(prefs));
          savePreferences(prefs);
          
          // Verify save by reloading
          const verifyPrefs = loadPreferences();
          console.log('[API] Verification - reloaded preferences:', JSON.stringify(verifyPrefs));
          
          res.writeHead(200, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type'
          });
          res.end(JSON.stringify({ 
            success: true, 
            message: 'Presets saved',
            saved: {
              presentation1: verifyPrefs.presentation1 || '',
              presentation2: verifyPrefs.presentation2 || '',
              presentation3: verifyPrefs.presentation3 || ''
            }
          }));
        } catch (error) {
          console.error('[API] Error saving presets:', error);
          console.error('[API] Error stack:', error.stack);
          res.writeHead(500, { 
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type'
          });
          res.end(JSON.stringify({ 
            error: error.message,
            code: error.code || 'UNKNOWN',
            details: process.platform === 'darwin' ? 'Check Console.app for detailed logs' : 'Check console output'
          }));
        }
      });
      return;
    }

    // POST /api/open-preset - Open a preset by name (1, 2, or 3)
    if (req.method === 'POST' && req.url === '/api/open-preset') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
      });
      
      req.on('end', async () => {
        try {
          const data = JSON.parse(body);
          const presetNumber = parseInt(data.preset, 10);
          
          if (isNaN(presetNumber) || presetNumber < 1 || presetNumber > 3) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'Preset must be 1, 2, or 3' }));
            return;
          }
          
          const prefs = loadPreferences();
          const presetKey = `presentation${presetNumber}`;
          const url = prefs[presetKey];
          
          if (!url) {
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: `Preset ${presetNumber} is not configured` }));
            return;
          }
          
          // Forward to open-presentation endpoint logic
          // We'll reuse the same code path
          console.log(`[API] Opening preset ${presetNumber}: ${url}`);
          
          // Close any existing presentation windows
          try {
            if (notesWindow && !notesWindow.isDestroyed()) {
              console.log('[API] Closing existing notes window');
              notesWindow.removeAllListeners('closed');
              notesWindow.close();
              notesWindow = null;
            }
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Closing existing presentation window');
              presentationWindow.removeAllListeners('closed');
              presentationWindow.close();
              presentationWindow = null;
            }
            currentSlide = null;
          } catch (error) {
            console.error('[API] Error closing existing windows:', error.message);
          }
          
          // Load preferences for monitor selection
          const displays = screen.getAllDisplays();
          const presentationDisplayId = Number(prefs.presentationDisplayId);
          const notesDisplayId = Number(prefs.notesDisplayId);
          const presentationDisplay = displays.find(d => d.id === presentationDisplayId) || displays[0];
          const notesDisplay = displays.find(d => d.id === notesDisplayId) || displays[0];
          
          console.log('[API] Using presentation display:', presentationDisplay.id);
          console.log('[API] Using notes display:', notesDisplay.id);
          
          // Create the presentation window (reuse open-presentation logic)
          presentationWindow = new BrowserWindow({
            x: presentationDisplay.bounds.x,
            y: presentationDisplay.bounds.y,
            width: presentationDisplay.bounds.width,
            height: presentationDisplay.bounds.height,
            fullscreen: true,
            frame: false,
            webPreferences: {
              nodeIntegration: false,
              contextIsolation: true,
              partition: GOOGLE_SESSION_PARTITION
            }
          });
          
          // Set up window handlers (same as open-presentation)
          presentationWindow.webContents.setWindowOpenHandler(({ url, frameName, features }) => {
            const windowOptions = {
              frame: false,
              webPreferences: {
                nodeIntegration: false,
                contextIsolation: true,
                partition: GOOGLE_SESSION_PARTITION
              }
            };
            // Use default size from Google Slides (no size/position override)
            return { action: 'allow', overrideBrowserWindowOptions: windowOptions };
          });
          
          const windowCreatedListener = (event, window) => {
            if (window !== presentationWindow && window !== mainWindow) {
              notesWindow = window;
              window.webContents.on('before-input-event', (event, input) => {
                if (input.key === 'Escape' && input.type === 'keyDown') {
                  event.preventDefault();
                  if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
                  if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
                }
              });
              // Set speaker notes window to fullscreen when it loads
              window.webContents.once('did-finish-load', () => {
                setSpeakerNotesFullscreen(window);
              });
              
              // Also try when DOM is ready (in case did-finish-load fires too early)
              window.webContents.once('dom-ready', () => {
                setTimeout(() => {
                  setSpeakerNotesFullscreen(window);
                }, 500);
              });
              
              app.removeListener('browser-window-created', windowCreatedListener);
            }
          };
          app.on('browser-window-created', windowCreatedListener);
          
          let sKeyPressed = false;
          const navigationListener = async (event, navUrl) => {
            const isPresentMode = (navUrl.includes('/present/') || navUrl.includes('localpresent')) && !navUrl.includes('/presentation/');
            if (isPresentMode && !sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
              sKeyPressed = true;
              await new Promise(resolve => setTimeout(resolve, 300));
              if (presentationWindow && !presentationWindow.isDestroyed()) {
                presentationWindow.focus();
                await new Promise(resolve => setTimeout(resolve, 50));
                presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
                presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
                presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
                presentationWindow.webContents.removeListener('did-navigate', navigationListener);
              }
            }
          };
          presentationWindow.webContents.on('did-navigate', navigationListener);
          
          presentationWindow.webContents.once('did-finish-load', async () => {
            if (!presentationWindow || presentationWindow.isDestroyed()) return;
            await new Promise(resolve => setTimeout(resolve, 200));
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 50));
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
            }
          });
          
          setTimeout(async () => {
            if (!sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
              sKeyPressed = true;
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 50));
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
              presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
              if (presentationWindow && !presentationWindow.isDestroyed()) {
                presentationWindow.webContents.removeListener('did-navigate', navigationListener);
              }
            }
          }, 1000);
          
          presentationWindow.on('closed', () => {
            presentationWindow = null;
            currentSlide = null;
          });
          
          presentationWindow.webContents.on('before-input-event', (event, input) => {
            if (input.key === 'Escape' && input.type === 'keyDown') {
              event.preventDefault();
              if (notesWindow && !notesWindow.isDestroyed()) notesWindow.close();
              if (presentationWindow && !presentationWindow.isDestroyed()) presentationWindow.close();
            }
          });
          
          const presentUrl = toPresentUrl(url);
          console.log('[API] Loading PRESENT URL:', presentUrl);
          lastPresentationUrl = url;
          currentSlide = 1;
          presentationWindow.loadURL(presentUrl);
          presentationWindow.show();
          
          // Ensure fullscreen on macOS
          presentationWindow.once('ready-to-show', () => {
            if (process.platform === 'darwin' && presentationWindow && !presentationWindow.isDestroyed()) {
              presentationWindow.setBounds({
                x: presentationDisplay.bounds.x,
                y: presentationDisplay.bounds.y,
                width: presentationDisplay.bounds.width,
                height: presentationDisplay.bounds.height
              });
              setTimeout(() => {
                if (presentationWindow && !presentationWindow.isDestroyed()) {
                  presentationWindow.setFullScreen(true);
                }
              }, 50);
            }
          });
          
          if (process.platform === 'darwin') {
            setTimeout(() => {
              if (presentationWindow && !presentationWindow.isDestroyed() && !presentationWindow.isFullScreen()) {
                presentationWindow.setBounds({
                  x: presentationDisplay.bounds.x,
                  y: presentationDisplay.bounds.y,
                  width: presentationDisplay.bounds.width,
                  height: presentationDisplay.bounds.height
                });
                presentationWindow.setFullScreen(true);
              }
            }, 200);
          }
          
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: true, message: `Preset ${presetNumber} opened`, url: url }));
        } catch (error) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: error.message }));
        }
      });
      return;
    }
    
    // 404 for unknown endpoints
    res.writeHead(404, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Not found' }));
  });
  
  const prefs = loadPreferences();
  const apiPort = prefs.apiPort || DEFAULT_API_PORT;
  
  httpServer.listen(apiPort, '0.0.0.0', () => {
    console.log(`[API] HTTP server listening on http://0.0.0.0:${apiPort}`);
  }).on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`[API] Port ${apiPort} is already in use`);
      dialog.showErrorBox(
        'Port Already in Use',
        `Port ${apiPort} is already in use. Another instance of Google Slides Opener may be running.\n\nPlease quit the other instance or change the API port in settings.`
      );
      // Don't exit the app, but the server won't start
    } else {
      console.error('[API] Server error:', err);
      dialog.showErrorBox(
        'Server Error',
        `Failed to start API server: ${err.message}`
      );
    }
  });
}

// Start web UI server for preset management
function startWebUiServer() {
  webUiServer = http.createServer((req, res) => {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    
    if (req.method === 'OPTIONS') {
      res.writeHead(200);
      res.end();
      return;
    }
    
    // GET / - Serve the web UI
    if (req.method === 'GET' && req.url === '/') {
      // Get configured API port for the web UI
      const prefs = loadPreferences();
      const apiPort = prefs.apiPort || DEFAULT_API_PORT;
      const webUiPort = prefs.webUiPort || DEFAULT_WEB_UI_PORT;
      
      // Get version and build number
      const packageJson = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
      const version = packageJson.version || '1.4.5';
      const buildNumber = packageJson.buildNumber || '1';
      const versionString = `v${version}.${buildNumber}`;
      
      // Get machine name or fallback to hostname
      // Escape HTML to prevent XSS
      const machineName = (prefs.machineName || os.hostname())
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
      
      const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Google Slides Opener - Preset Manager</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    @media (max-width: 768px) {
      body {
        padding: 8px;
        align-items: flex-start;
      }
    }
    .container {
      background: white;
      border-radius: 16px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3);
      max-width: 600px;
      width: 100%;
      padding: 40px;
      transition: all 0.3s;
    }
    @media (max-width: 768px) {
      .container {
        padding: 16px 20px;
      }
    }
    body.notes-visible .container {
      max-width: 85%;
    }
    h1 {
      color: #333;
      margin-top: 0;
      margin-bottom: 0;
      padding-top: 8px;
      padding-bottom: 8px;
      font-size: 28px;
      transition: all 0.3s;
      display: flex;
      align-items: center;
      gap: 12px;
    }
    body.notes-visible h1 {
      font-size: 20px;
      padding-top: 4px;
      padding-bottom: 4px;
    }
    .system-icon {
      width: 32px;
      height: 32px;
      color: #667eea;
      flex-shrink: 0;
    }
    body.notes-visible .system-icon {
      width: 24px;
      height: 24px;
    }
    .preset-group {
      margin-bottom: 24px;
    }
    label {
      display: block;
      font-weight: 600;
      color: #333;
      margin-bottom: 8px;
      font-size: 14px;
    }
    input[type="text"] {
      width: 100%;
      padding: 12px 16px;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      font-size: 14px;
      transition: border-color 0.2s;
    }
    input[type="text"]:focus {
      outline: none;
      border-color: #667eea;
    }
    .btn {
      width: 100%;
      padding: 14px;
      background: #667eea;
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.2s;
      margin-top: 8px;
    }
    .btn:hover {
      background: #5568d3;
    }
    .btn:active {
      transform: scale(0.98);
    }
    .btn-secondary {
      background: #6c757d;
      margin-top: 12px;
    }
    .btn-secondary:hover {
      background: #5a6268;
    }
    .status {
      position: fixed;
      bottom: 24px;
      left: 50%;
      transform: translateX(-50%) translateY(100px);
      padding: 12px 20px;
      border-radius: 8px;
      text-align: center;
      font-size: 14px;
      font-weight: 500;
      opacity: 0;
      pointer-events: none;
      z-index: 1000;
      min-width: 200px;
      max-width: 90%;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      transition: all 0.3s ease;
    }
    .status.success {
      background: #d4edda;
      color: #155724;
      border: 1px solid #c3e6cb;
      opacity: 1;
      transform: translateX(-50%) translateY(0);
      pointer-events: auto;
    }
    .status.error {
      background: #f8d7da;
      color: #721c24;
      border: 1px solid #f5c6cb;
      opacity: 1;
      transform: translateX(-50%) translateY(0);
      pointer-events: auto;
    }
    .info {
      background: #e7f3ff;
      border: 1px solid #b3d9ff;
      color: #004085;
      padding: 12px;
      border-radius: 8px;
      margin-bottom: 24px;
      font-size: 13px;
      line-height: 1.5;
    }
    .controls-section {
      margin-bottom: 30px;
      padding-top: 20px;
      border-top: 2px solid #e0e0e0;
    }
    .controls-section h3 {
      color: #333;
      font-size: 18px;
      margin-bottom: 12px;
      margin-top: 20px;
    }
    .controls-section h3:first-child {
      margin-top: 0;
    }
    .controls-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 10px;
      margin-bottom: 20px;
    }
    .btn-control {
      padding: 12px 16px;
      background: #f8f9fa;
      color: #333;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
    }
    .btn-control:hover {
      background: #667eea;
      color: white;
      border-color: #667eea;
      transform: translateY(-2px);
      box-shadow: 0 4px 8px rgba(102, 126, 234, 0.3);
    }
    .btn-control:active {
      transform: translateY(0);
    }
    .btn-control:disabled {
      opacity: 0.5;
      cursor: not-allowed;
      transform: none;
    }
    .btn-icon {
      width: 18px;
      height: 18px;
      stroke-width: 2.5;
    }
    .tabs {
      display: flex;
      gap: 8px;
      margin-bottom: 24px;
      border-bottom: 2px solid #e0e0e0;
      transition: all 0.3s;
    }
    body.notes-visible .tabs {
      margin-bottom: 12px;
      border-bottom-width: 1px;
    }
    .tab-btn {
      padding: 12px 24px;
      background: transparent;
      border: none;
      border-bottom: 3px solid transparent;
      font-size: 16px;
      font-weight: 600;
      color: #666;
      cursor: pointer;
      transition: all 0.3s;
      margin-bottom: -2px;
    }
    body.notes-visible .tab-btn {
      padding: 8px 16px;
      font-size: 13px;
      border-bottom-width: 2px;
    }
    .tab-btn:hover {
      color: #333;
      background: #f8f9fa;
    }
    .tab-btn.active {
      color: #667eea;
      border-bottom-color: #667eea;
    }
    .tab-content {
      display: none;
    }
    .tab-content.active {
      display: block;
    }
    /* Floating tooltips that don't affect layout */
    .btn-control[data-tooltip] {
      position: relative;
    }
    .btn-control[data-tooltip]:hover::after {
      content: attr(data-tooltip);
      position: absolute;
      bottom: calc(100% + 8px);
      left: 50%;
      transform: translateX(-50%);
      padding: 6px 10px;
      background: #333;
      color: white;
      font-size: 12px;
      font-weight: normal;
      white-space: nowrap;
      border-radius: 4px;
      pointer-events: none;
      z-index: 1000;
      box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    }
    .btn-control[data-tooltip]:hover::before {
      content: '';
      position: absolute;
      bottom: calc(100% + 2px);
      left: 50%;
      transform: translateX(-50%);
      border: 5px solid transparent;
      border-top-color: #333;
      pointer-events: none;
      z-index: 1001;
    }
    /* Build number display */
    .build-number {
      position: fixed;
      bottom: 8px;
      left: 8px;
      font-size: 11px;
      color: #999;
      opacity: 0.7;
      z-index: 10;
    }
    /* Remote tab - big buttons for mobile */
    .remote-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 20px;
      padding-bottom: 12px;
      border-bottom: 2px solid #e0e0e0;
    }
    .remote-header h2 {
      margin: 0;
      font-size: 20px;
      color: #333;
    }
    .notes-toggle-btn {
      background: #667eea;
      color: white;
      border: none;
      border-radius: 8px;
      padding: 10px 14px;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 14px;
      transition: all 0.2s;
    }
    .notes-toggle-btn:hover {
      background: #5568d3;
    }
    .notes-toggle-btn.active {
      background: #764ba2;
    }
    .notes-toggle-btn svg {
      width: 18px;
      height: 18px;
    }
    .remote-controls {
      display: flex;
      flex-direction: column;
      gap: 20px;
      padding: 20px 0;
      transition: all 0.3s;
    }
    .remote-controls.with-notes {
      gap: 12px;
    }
    .remote-btn {
      width: 100%;
      padding: 40px 20px;
      font-size: 24px;
      font-weight: 700;
      border-radius: 12px;
      border: none;
      cursor: pointer;
      transition: all 0.3s;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 12px;
      min-height: 120px;
    }
    .remote-controls.with-notes .remote-btn {
      padding: 20px 16px;
      font-size: 18px;
      min-height: 70px;
    }
    .remote-btn-prev {
      background: #667eea;
      color: white;
    }
    .remote-btn-prev:hover {
      background: #5568d3;
      transform: scale(1.02);
    }
    .remote-btn-next {
      background: #667eea;
      color: white;
    }
    .remote-btn-next:hover {
      background: #5568d3;
      transform: scale(1.02);
    }
    .remote-btn:active {
      transform: scale(0.98);
    }
    .remote-btn svg {
      width: 32px;
      height: 32px;
      transition: all 0.3s;
    }
    .remote-controls.with-notes .remote-btn svg {
      width: 24px;
      height: 24px;
    }
    /* Speaker notes display */
    .speaker-notes-container {
      display: none;
      margin-top: 12px;
      transition: all 0.3s;
    }
    .speaker-notes-container.visible {
      display: block;
    }
    .speaker-notes-content-wrapper {
      background: #f8f9fa;
      border: 2px solid #e0e0e0;
      border-radius: 12px;
      padding: 16px;
      height: 400px;
      overflow-y: auto;
      overflow-x: hidden;
    }
    .speaker-notes-content {
      color: #333;
      font-size: 18px;
      line-height: 1.6;
      white-space: pre-wrap;
      word-wrap: break-word;
    }
    .speaker-notes-content.zoom-small {
      font-size: 16px;
    }
    .speaker-notes-content.zoom-large {
      font-size: 22px;
    }
    .notes-zoom-controls {
      display: none;
      justify-content: center;
      gap: 10px;
      margin-top: 12px;
      position: sticky;
      bottom: 0;
      background: white;
      padding: 8px 0;
      z-index: 10;
    }
    .notes-zoom-controls.visible {
      display: flex;
    }
    .notes-zoom-btn {
      background: #f8f9fa;
      border: 2px solid #e0e0e0;
      border-radius: 6px;
      padding: 8px 16px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      color: #333;
      transition: all 0.2s;
    }
    .notes-zoom-btn:hover {
      background: #667eea;
      color: white;
      border-color: #667eea;
    }
    /* Bigger slide control buttons */
    .btn-control-large {
      padding: 20px 24px;
      font-size: 18px;
      min-height: 60px;
    }
    /* Stagetimer display */
    .stagetimer-container {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      border-radius: 12px;
      padding: 16px 20px;
      padding-bottom: 16px;
      margin-bottom: 20px;
      text-align: center;
      color: white;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      position: relative;
      height: 160px;
      overflow: visible;
    }
    .stagetimer-container.error {
      background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%);
    }
    .stagetimer-container.disabled {
      background: #e0e0e0;
      color: #666;
    }
    .stagetimer-label {
      font-size: 16px;
      opacity: 0.95;
      margin-bottom: 6px;
      font-weight: 600;
    }
    .stagetimer-time {
      font-size: 42px;
      font-weight: 700;
      font-variant-numeric: tabular-nums;
      letter-spacing: 2px;
      margin: 4px 0;
      line-height: 1.2;
    }
    .stagetimer-status {
      font-size: 14px;
      opacity: 0.9;
      margin-top: 6px;
      font-weight: 500;
    }
    .stagetimer-name {
      display: none;
    }
    .stagetimer-container.running {
      background: linear-gradient(135deg, #4caf50 0%, #388e3c 100%);
    }
    .stagetimer-container.warning {
      background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%);
    }
    .stagetimer-container.critical {
      background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%);
    }
    /* Stagetimer messages - absolutely positioned to prevent layout shift */
    .stagetimer-messages {
      position: absolute;
      bottom: 0;
      left: 0;
      right: 0;
      margin: 0;
      padding: 12px 24px 16px 24px;
      border-top: 1px solid rgba(255, 255, 255, 0.3);
      background: linear-gradient(to top, rgba(0, 0, 0, 0.4) 0%, rgba(0, 0, 0, 0.2) 50%, transparent 100%);
      border-radius: 0 0 12px 12px;
      max-height: 100px;
      overflow-y: auto;
      overflow-x: hidden;
      display: none;
      backdrop-filter: blur(8px);
      z-index: 10;
    }
    .stagetimer-messages.visible {
      display: block;
    }
    .stagetimer-message {
      background: rgba(255, 255, 255, 0.15);
      border-radius: 8px;
      padding: 12px 16px;
      margin-bottom: 8px;
      font-size: 14px;
      line-height: 1.5;
      backdrop-filter: blur(10px);
    }
    .stagetimer-message:last-child {
      margin-bottom: 0;
    }
    .stagetimer-message.white {
      background: rgba(255, 255, 255, 0.2);
      color: white;
    }
    .stagetimer-message.green {
      background: rgba(76, 175, 80, 0.3);
      color: #c8e6c9;
    }
    .stagetimer-message.red {
      background: rgba(244, 67, 54, 0.3);
      color: #ffcdd2;
    }
    .stagetimer-message.bold {
      font-weight: 700;
    }
    .stagetimer-message.uppercase {
      text-transform: uppercase;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>
      <svg class="system-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <rect x="2" y="4" width="20" height="12" rx="2" ry="2"></rect>
        <line x1="6" y1="20" x2="18" y2="20"></line>
        <line x1="8" y1="16" x2="8" y2="20"></line>
        <line x1="16" y1="16" x2="16" y2="20"></line>
        <circle cx="12" cy="10" r="3" fill="currentColor"></circle>
        <polygon points="10 10 12 9 14 10 12 11" fill="white"></polygon>
      </svg>
      ${machineName}
    </h1>
    
    <!-- Tabs -->
    <div class="tabs">
      <button class="tab-btn active" data-tab="remote">Remote</button>
      <button class="tab-btn" data-tab="controls">Controls</button>
      <button class="tab-btn" data-tab="settings">Settings</button>
    </div>
    
    <!-- Remote Tab (Default) -->
    <div id="tab-remote" class="tab-content active">
      <div class="remote-header">
        <h2>Remote Control</h2>
        <button type="button" class="notes-toggle-btn" id="notes-toggle-btn" title="Toggle speaker notes">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
          </svg>
          Notes
        </button>
      </div>
      <div class="stagetimer-container disabled" id="stagetimer-container" style="display: none;">
        <div class="stagetimer-label" id="stagetimer-label">Stage Timer</div>
        <div class="stagetimer-time" id="stagetimer-time">--:--</div>
        <div class="stagetimer-status" id="stagetimer-status">Not configured</div>
        <div class="stagetimer-messages" id="stagetimer-messages"></div>
      </div>
      <div class="remote-controls" id="remote-controls">
        <button type="button" class="remote-btn remote-btn-prev" id="remote-btn-prev">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3">
            <polyline points="15 18 9 12 15 6"></polyline>
          </svg>
          Previous Slide
        </button>
        <button type="button" class="remote-btn remote-btn-next" id="remote-btn-next">
          Next Slide
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3">
            <polyline points="9 18 15 12 9 6"></polyline>
          </svg>
        </button>
      </div>
      <div class="speaker-notes-container" id="speaker-notes-container">
        <div class="notes-zoom-controls" id="notes-zoom-controls">
          <button type="button" class="notes-zoom-btn" id="notes-zoom-out">Zoom Out</button>
          <button type="button" class="notes-zoom-btn" id="notes-zoom-in">Zoom In</button>
        </div>
        <div class="speaker-notes-content-wrapper">
          <div class="speaker-notes-content" id="speaker-notes-content">Loading notes...</div>
        </div>
      </div>
    </div>
    
    <!-- Controls Tab -->
    <div id="tab-controls" class="tab-content">
      <div class="info">
        Use these controls to manage your active presentation.
      </div>
      
      <!-- Open Presentation -->
      <div class="controls-section">
        <h3>Open Presentation</h3>
        <div class="preset-group">
          <label for="presentation-url">Google Slides URL</label>
          <input type="text" id="presentation-url" name="presentation-url" placeholder="https://docs.google.com/presentation/d/..." />
        </div>
        <div style="display: flex; gap: 10px;">
          <button type="button" class="btn" id="btn-open-presentation" style="flex: 1;">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;">
              <polyline points="5 12 3 12 12 3 21 12 19 12"></polyline>
              <path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path>
              <polyline points="9 21 9 12 15 12 15 21"></polyline>
            </svg>
            Launch Presentation
          </button>
          <button type="button" class="btn" id="btn-open-presentation-with-notes" style="flex: 1;">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;">
              <polyline points="5 12 3 12 12 3 21 12 19 12"></polyline>
              <path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path>
              <polyline points="9 21 9 12 15 12 15 21"></polyline>
            </svg>
            Launch with Notes
          </button>
        </div>
      </div>
      
      <!-- Preset Presentations -->
      <div class="controls-section">
        <h3>Preset Presentations</h3>
        <div id="preset-buttons-container" style="display: flex; flex-direction: column; gap: 10px;">
          <!-- Preset buttons will be dynamically loaded here -->
        </div>
      </div>
      
      <!-- Speaker Notes Controls -->
      <div class="controls-section">
        <h3>Speaker Notes</h3>
        <button type="button" class="btn-control" id="btn-start-notes" title="Start speaker notes window">
          <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="16" y1="13" x2="8" y2="13"></line>
            <line x1="16" y1="17" x2="8" y2="17"></line>
            <polyline points="10 9 9 9 8 9"></polyline>
          </svg>
          Start Notes
        </button>
      </div>
      
      <!-- Presentation Controls -->
      <div class="controls-section">
        <h3>Presentation Controls</h3>
        <div class="controls-grid">
          <button type="button" class="btn-control" id="btn-reload" data-tooltip="Reload presentation and return to current slide">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="23 4 23 10 17 10"></polyline>
              <polyline points="1 20 1 14 7 14"></polyline>
              <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path>
            </svg>
            Reload Presentation
          </button>
          <button type="button" class="btn-control" id="btn-close-presentation" data-tooltip="Close current presentation">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <line x1="18" y1="6" x2="6" y2="18"></line>
              <line x1="6" y1="6" x2="18" y2="18"></line>
            </svg>
            Close Presentation
          </button>
        </div>
      </div>
    </div>
    
    <!-- Settings Tab (Hidden by default) -->
    <div id="tab-settings" class="tab-content">
      <div class="info">
        Configure preset presentations. These can be opened from Companion using "Open Presentation 1", "Open Presentation 2", or "Open Presentation 3" actions.
      </div>
      
      <form id="preset-form">
      <div class="preset-group">
        <label for="preset1">Presentation 1</label>
        <input type="text" id="preset1" name="preset1" placeholder="https://docs.google.com/presentation/d/..." />
      </div>
      
      <div class="preset-group">
        <label for="preset2">Presentation 2</label>
        <input type="text" id="preset2" name="preset2" placeholder="https://docs.google.com/presentation/d/..." />
      </div>
      
      <div class="preset-group">
        <label for="preset3">Presentation 3</label>
        <input type="text" id="preset3" name="preset3" placeholder="https://docs.google.com/presentation/d/..." />
      </div>
      
        <button type="submit" class="btn">Save Presets</button>
        <button type="button" class="btn btn-secondary" id="load-btn">Load Current Presets</button>
      </form>
      
      <!-- Stagetimer Integration -->
      <div class="controls-section" style="margin-top: 40px;">
        <h3>Stagetimer.io Integration</h3>
        <div class="info" style="margin-bottom: 20px;">
          Connect to your stagetimer.io room to display live timer data. Get your Room ID and API Key from the stagetimer.io controller page.
        </div>
        <div class="preset-group">
          <label for="stagetimer-room-id">Room ID</label>
          <input type="text" id="stagetimer-room-id" name="stagetimer-room-id" placeholder="Enter your stagetimer.io Room ID" />
        </div>
        <div class="preset-group">
          <label for="stagetimer-api-key">API Key</label>
          <input type="password" id="stagetimer-api-key" name="stagetimer-api-key" placeholder="Enter your stagetimer.io API Key" />
        </div>
        <div style="display: flex; align-items: center; gap: 10px; margin-top: 12px;">
          <input type="checkbox" id="stagetimer-enabled" style="width: auto;" />
          <label for="stagetimer-enabled" style="margin: 0; font-weight: normal;">Enable timer display</label>
        </div>
        <div style="display: flex; align-items: center; gap: 10px; margin-top: 12px;">
          <input type="checkbox" id="stagetimer-visible" style="width: auto;" checked />
          <label for="stagetimer-visible" style="margin: 0; font-weight: normal;">Show timer on Remote tab</label>
        </div>
        <button type="button" class="btn" id="btn-save-stagetimer" style="margin-top: 12px;">Save Stagetimer Settings</button>
        <button type="button" class="btn btn-secondary" id="btn-load-stagetimer" style="margin-top: 8px;">Load Current Settings</button>
      </div>
      
      <!-- Debug Console -->
      <div class="controls-section" style="margin-top: 40px;">
        <h3>Debug Console</h3>
        <div class="info" style="margin-bottom: 10px;">
          Console output for debugging stagetimer integration and other issues.
        </div>
        <div style="background: #1e1e1e; color: #d4d4d4; padding: 12px; border-radius: 8px; font-family: 'Courier New', monospace; font-size: 12px; max-height: 300px; overflow-y: auto; margin-bottom: 10px;" id="debug-console">
          <div style="color: #888;">Console ready. Logs will appear here...</div>
        </div>
        <button type="button" class="btn btn-secondary" id="btn-clear-console" style="margin-top: 8px;">Clear Console</button>
      </div>
    </div>
    
    <div id="status" class="status"></div>
    <div class="build-number">${versionString}</div>
  </div>
  
  <script>
    const form = document.getElementById('preset-form');
    const loadBtn = document.getElementById('load-btn');
    const status = document.getElementById('status');
    // Use relative URLs so API calls go through the Web UI server (port 80)
    // The Web UI server will proxy these requests to the API server (port 9595)
    // This allows the Web UI to work even when only port 80 is accessible from the network
    const API_BASE = '';
    
    // Debug: Log the API base URL for troubleshooting
    console.log('[Web UI] Using relative API URLs (proxied through Web UI server on port 80)');
    console.log('[Web UI] window.location.hostname:', window.location.hostname);
    console.log('[Web UI] window.location.host:', window.location.host);
    
    function showStatus(message, isError) {
      status.textContent = message;
      status.className = 'status ' + (isError ? 'error' : 'success');
      setTimeout(() => {
        status.className = 'status';
        status.textContent = ''; // Clear text when hidden
      }, 3000);
    }
    
    // Prevent native tooltips and use custom floating ones (skip prev/next slide buttons)
    document.querySelectorAll('.btn-control[title]').forEach(btn => {
      // Skip prev/next slide buttons - no tooltips for those
      if (btn.id === 'btn-prev-slide' || btn.id === 'btn-next-slide') {
        btn.removeAttribute('title');
        return;
      }
      
      const titleText = btn.getAttribute('title');
      btn.setAttribute('data-tooltip', titleText);
      btn.removeAttribute('title'); // Remove native title to prevent layout shift
      
      // Restore title for accessibility when not hovering
      btn.addEventListener('mouseenter', function() {
        this.removeAttribute('title');
      });
      btn.addEventListener('mouseleave', function() {
        this.setAttribute('title', titleText);
      });
    });
    
    // Tab switching
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const tabName = btn.getAttribute('data-tab');
        
        // Update buttons
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        
        // Update content
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        document.getElementById('tab-' + tabName).classList.add('active');
      });
    });
    
    function apiCall(endpoint, method = 'POST') {
      const url = API_BASE + endpoint;
      console.log('[Web UI] Making API call:', method, url);
      
      return fetch(url, {
        method: method,
        headers: { 'Content-Type': 'application/json' }
      })
        .then(res => {
          if (!res.ok) {
            throw new Error('HTTP error! status: ' + res.status);
          }
          return res.json();
        })
        .then(result => {
          console.log('[Web UI] API response:', result);
          if (result.success !== false) {
            showStatus(result.message || 'Action completed successfully', false);
          } else {
            showStatus(result.error || 'Action failed', true);
          }
          return result;
        })
        .catch(err => {
          console.error('[Web UI] API call error:', err);
          console.error('[Web UI] Failed URL:', url);
          let errorMsg = 'Failed: ' + err.message;
          if (err.message.includes('Failed to fetch') || err.message.includes('NetworkError')) {
            errorMsg += ' (Cannot reach API server at ' + API_BASE + '. Check network connection and firewall settings.)';
          } else {
            errorMsg += ' (Make sure the app is running)';
          }
          showStatus(errorMsg, true);
          throw err;
        });
    }
    
    // Haptic feedback function for mobile devices
    function triggerHapticFeedback() {
      if ('vibrate' in navigator) {
        // Light vibration for button press
        navigator.vibrate(10);
      }
    }
    
    // Set up control buttons
    document.getElementById('btn-reload').addEventListener('click', () => {
      apiCall('/api/reload-presentation');
    });
    
    document.getElementById('btn-close-presentation').addEventListener('click', () => {
      apiCall('/api/close-presentation');
    });
    
    // Remote tab buttons
    // Speaker notes functionality
    let notesVisible = false;
    let notesZoomLevel = 1; // Numeric zoom level (1 = normal, can go up/down continuously)
    
    function loadSpeakerNotes() {
      fetch(API_BASE + '/api/get-speaker-notes')
        .then(res => res.json())
        .then(data => {
          const notesContent = document.getElementById('speaker-notes-content');
          if (data.success && data.notes) {
            notesContent.textContent = data.notes;
          } else {
            notesContent.textContent = data.notes || 'No notes available. Make sure speaker notes are open.';
          }
        })
        .catch(err => {
          console.error('Failed to load speaker notes:', err);
          document.getElementById('speaker-notes-content').textContent = 'Failed to load notes.';
        });
    }
    
    function updateNotesZoom() {
      const notesContent = document.getElementById('speaker-notes-content');
      // Calculate font size based on zoom level (18px base, +/- 2px per level)
      const baseSize = 18;
      const fontSize = baseSize + ((notesZoomLevel - 1) * 2);
      notesContent.style.fontSize = fontSize + 'px';
    }
    
    document.getElementById('notes-toggle-btn').addEventListener('click', () => {
      const btn = document.getElementById('notes-toggle-btn');
      const container = document.getElementById('speaker-notes-container');
      const controls = document.getElementById('remote-controls');
      const zoomControls = document.getElementById('notes-zoom-controls');
      const body = document.body;
      
      if (!notesVisible) {
        // Opening notes - first check if speaker notes window is open, if not, open it
        fetch(API_BASE + '/api/get-speaker-notes')
          .then(res => res.json())
          .then(data => {
            // If notes window is not open, open it first
            if (!data.success && data.error && data.error.includes('No speaker notes window')) {
              console.log('[Web UI] Speaker notes not open, opening them first...');
              return apiCall('/api/open-speaker-notes').then(() => {
                // Wait a moment for notes to open, then show the UI
                setTimeout(() => {
                  notesVisible = true;
                  btn.classList.add('active');
                  container.classList.add('visible');
                  controls.classList.add('with-notes');
                  zoomControls.classList.add('visible');
                  body.classList.add('notes-visible');
                  loadSpeakerNotes();
                  // Refresh notes every 2 seconds when visible
                  if (window.notesRefreshInterval) clearInterval(window.notesRefreshInterval);
                  window.notesRefreshInterval = setInterval(loadSpeakerNotes, 2000);
                }, 1000);
              });
            } else {
              // Notes are already open, just show the UI
              notesVisible = true;
              btn.classList.add('active');
              container.classList.add('visible');
              controls.classList.add('with-notes');
              zoomControls.classList.add('visible');
              body.classList.add('notes-visible');
              loadSpeakerNotes();
              // Refresh notes every 2 seconds when visible
              if (window.notesRefreshInterval) clearInterval(window.notesRefreshInterval);
              window.notesRefreshInterval = setInterval(loadSpeakerNotes, 2000);
            }
          })
          .catch(err => {
            console.error('[Web UI] Error checking/opening speaker notes:', err);
            // Try to open notes anyway
            apiCall('/api/open-speaker-notes').then(() => {
              setTimeout(() => {
                notesVisible = true;
                btn.classList.add('active');
                container.classList.add('visible');
                controls.classList.add('with-notes');
                zoomControls.classList.add('visible');
                body.classList.add('notes-visible');
                loadSpeakerNotes();
                if (window.notesRefreshInterval) clearInterval(window.notesRefreshInterval);
                window.notesRefreshInterval = setInterval(loadSpeakerNotes, 2000);
              }, 1000);
            });
          });
      } else {
        // Closing notes
        notesVisible = false;
        btn.classList.remove('active');
        container.classList.remove('visible');
        controls.classList.remove('with-notes');
        zoomControls.classList.remove('visible');
        body.classList.remove('notes-visible');
        if (window.notesRefreshInterval) {
          clearInterval(window.notesRefreshInterval);
          window.notesRefreshInterval = null;
        }
      }
    });
    
    // Notes zoom controls
    // Allow continuous zoom - the API zooms the actual notes window, and we update Web UI display
    document.getElementById('notes-zoom-out').addEventListener('click', () => {
      // Decrease zoom level (minimum 0.5x)
      if (notesZoomLevel > 0.5) {
        notesZoomLevel = Math.max(0.5, notesZoomLevel - 0.5);
      }
      updateNotesZoom();
      apiCall('/api/zoom-out-notes').catch(() => {}); // Zoom the actual notes window
    });
    
    document.getElementById('notes-zoom-in').addEventListener('click', () => {
      // Increase zoom level (no maximum limit)
      notesZoomLevel += 0.5;
      updateNotesZoom();
      apiCall('/api/zoom-in-notes').catch(() => {}); // Zoom the actual notes window
    });
    
    document.getElementById('remote-btn-next').addEventListener('click', () => {
      triggerHapticFeedback();
      apiCall('/api/next-slide').then(() => {
        // Refresh notes after slide change
        if (notesVisible) {
          loadSpeakerNotes();
        }
      });
    });
    
    document.getElementById('remote-btn-prev').addEventListener('click', () => {
      triggerHapticFeedback();
      apiCall('/api/previous-slide').then(() => {
        // Refresh notes after slide change
        if (notesVisible) {
          loadSpeakerNotes();
        }
      });
    });
    
    // Helper function to validate and open presentation
    function openPresentation(url, withNotes = false) {
      if (!url) {
        showStatus('Please enter a Google Slides URL', true);
        document.getElementById('presentation-url').focus();
        return;
      }
      
      // Validate it looks like a Google Slides URL
      if (!url.includes('docs.google.com/presentation')) {
        showStatus('Please enter a valid Google Slides URL', true);
        document.getElementById('presentation-url').focus();
        return;
      }
      
      const endpoint = withNotes ? '/api/open-presentation-with-notes' : '/api/open-presentation';
      const btnId = withNotes ? 'btn-open-presentation-with-notes' : 'btn-open-presentation';
      const btn = document.getElementById(btnId);
      const originalText = btn.innerHTML;
      
      // Disable both buttons during request
      document.getElementById('btn-open-presentation').disabled = true;
      document.getElementById('btn-open-presentation-with-notes').disabled = true;
      btn.innerHTML = 'Opening...';
      
      fetch(API_BASE + endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url: url })
      })
        .then(res => {
          if (!res.ok) {
            return res.json().then(data => {
              throw new Error(data.error || 'HTTP error! status: ' + res.status);
            });
          }
          return res.json();
        })
        .then(result => {
          if (result.success) {
            showStatus(result.message || 'Presentation opened successfully!', false);
            document.getElementById('presentation-url').value = ''; // Clear the input
          } else {
            showStatus('Failed to open: ' + (result.error || 'Unknown error'), true);
          }
        })
        .catch(err => {
          console.error('Open presentation error:', err);
          showStatus('Failed to open presentation: ' + err.message + ' (Make sure the app is running)', true);
        })
        .finally(() => {
          document.getElementById('btn-open-presentation').disabled = false;
          document.getElementById('btn-open-presentation-with-notes').disabled = false;
          document.getElementById('btn-open-presentation').innerHTML = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;"><polyline points="5 12 3 12 12 3 21 12 19 12"></polyline><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path><polyline points="9 21 9 12 15 12 15 21"></polyline></svg>Launch Presentation';
          document.getElementById('btn-open-presentation-with-notes').innerHTML = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;"><polyline points="5 12 3 12 12 3 21 12 19 12"></polyline><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path><polyline points="9 21 9 12 15 12 15 21"></polyline></svg>Launch with Notes';
        });
    }
    
    // Open presentation button (without notes)
    document.getElementById('btn-open-presentation').addEventListener('click', () => {
      const url = document.getElementById('presentation-url').value.trim();
      openPresentation(url, false);
    });
    
    // Open presentation with notes button
    document.getElementById('btn-open-presentation-with-notes').addEventListener('click', () => {
      const url = document.getElementById('presentation-url').value.trim();
      openPresentation(url, true);
    });
    
    // Start notes button
    document.getElementById('btn-start-notes').addEventListener('click', () => {
      apiCall('/api/open-speaker-notes');
    });
    
    // Allow Enter key to trigger open (without notes)
    document.getElementById('presentation-url').addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        document.getElementById('btn-open-presentation').click();
      }
    });
    
    // Speaker notes controls removed from default Controls tab - moved to Settings if needed later
    
    // Function to create preset buttons
    function createPresetButtons(presets) {
      const container = document.getElementById('preset-buttons-container');
      container.innerHTML = '';
      
      for (let i = 1; i <= 3; i++) {
        const presetUrl = presets[\`presentation\${i}\`];
        if (!presetUrl || presetUrl.trim() === '') {
          continue; // Skip empty presets
        }
        
        const presetGroup = document.createElement('div');
        presetGroup.style.cssText = 'display: flex; gap: 10px; margin-bottom: 10px;';
        
        const label = document.createElement('div');
        label.textContent = \`Presentation \${i}:\`;
        label.style.cssText = 'font-weight: 600; color: #333; padding: 12px 0; min-width: 120px; font-size: 14px;';
        
        const buttonGroup = document.createElement('div');
        buttonGroup.style.cssText = 'display: flex; gap: 10px; flex: 1;';
        
        const launchBtn = document.createElement('button');
        launchBtn.type = 'button';
        launchBtn.className = 'btn';
        launchBtn.style.cssText = 'flex: 1;';
        launchBtn.innerHTML = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;"><polyline points="5 12 3 12 12 3 21 12 19 12"></polyline><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path><polyline points="9 21 9 12 15 12 15 21"></polyline></svg>Launch';
        launchBtn.addEventListener('click', () => {
          openPresentation(presetUrl, false);
        });
        
        const launchWithNotesBtn = document.createElement('button');
        launchWithNotesBtn.type = 'button';
        launchWithNotesBtn.className = 'btn';
        launchWithNotesBtn.style.cssText = 'flex: 1;';
        launchWithNotesBtn.innerHTML = '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 18px; height: 18px; display: inline-block; vertical-align: middle; margin-right: 8px;"><polyline points="5 12 3 12 12 3 21 12 19 12"></polyline><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"></path><polyline points="9 21 9 12 15 12 15 21"></polyline></svg>Launch with Notes';
        launchWithNotesBtn.addEventListener('click', () => {
          openPresentation(presetUrl, true);
        });
        
        buttonGroup.appendChild(launchBtn);
        buttonGroup.appendChild(launchWithNotesBtn);
        
        presetGroup.appendChild(label);
        presetGroup.appendChild(buttonGroup);
        container.appendChild(presetGroup);
      }
      
      if (container.children.length === 0) {
        container.innerHTML = '<div style="color: #999; font-style: italic; padding: 20px; text-align: center;">No preset presentations configured. Go to Settings to add presets.</div>';
      }
    }
    
    // Test API connection on page load
    fetch(API_BASE + '/api/status')
      .then(res => {
        if (!res.ok) {
          throw new Error('HTTP error! status: ' + res.status);
        }
        return res.json();
      })
      .then(data => {
        console.log('[Web UI] API connection successful:', data);
        // API is reachable, now load presets
        return fetch(API_BASE + '/api/presets');
      })
      .then(res => {
        if (!res.ok) {
          throw new Error('HTTP error! status: ' + res.status);
        }
        return res.json();
      })
      .then(data => {
        document.getElementById('preset1').value = data.presentation1 || '';
        document.getElementById('preset2').value = data.presentation2 || '';
        document.getElementById('preset3').value = data.presentation3 || '';
        // Create preset buttons in Controls tab
        createPresetButtons(data);
      })
      .catch(err => {
        console.error('[Web UI] Failed to connect to API:', err);
        console.error('[Web UI] API_BASE was:', API_BASE);
        // Show a warning if API is not reachable
        const statusEl = document.getElementById('status');
        if (statusEl) {
          statusEl.textContent = 'Warning: Cannot connect to API server at ' + API_BASE + '. Controls may not work.';
          statusEl.className = 'status error';
          setTimeout(() => {
            statusEl.className = 'status';
            statusEl.textContent = '';
          }, 10000);
        }
      });
    
    // Load button
    loadBtn.addEventListener('click', () => {
      fetch(API_BASE + '/api/presets')
        .then(res => {
          if (!res.ok) {
            throw new Error('HTTP error! status: ' + res.status);
          }
          return res.json();
        })
        .then(data => {
          document.getElementById('preset1').value = data.presentation1 || '';
          document.getElementById('preset2').value = data.presentation2 || '';
          document.getElementById('preset3').value = data.presentation3 || '';
          showStatus('Presets loaded', false);
          // Update preset buttons in Controls tab
          createPresetButtons(data);
        })
        .catch(err => {
          console.error('Load error:', err);
          showStatus('Failed to load presets: ' + err.message + ' (Make sure the app is running)', true);
        });
    });
    
    // Stagetimer integration
    let stagetimerPollInterval = null;
    let stagetimerTimerInfoInterval = null;
    let stagetimerDisplayInterval = null; // For local time updates
    let stagetimerEnabled = false;
    let stagetimerVisible = true;
    let stagetimerState = null; // Store timer state for local calculation
    let stagetimerBackoffDelay = 10000; // Start with 10 second delay, increases on 429 errors
    let stagetimerIsBackingOff = false;
    
    function loadStagetimerSettings() {
      fetch(API_BASE + '/api/stagetimer-settings')
        .then(res => res.json())
        .then(data => {
          document.getElementById('stagetimer-room-id').value = data.roomId || '';
          document.getElementById('stagetimer-api-key').value = data.apiKey || '';
          document.getElementById('stagetimer-enabled').checked = data.enabled !== false;
          document.getElementById('stagetimer-visible').checked = data.visible !== false;
          stagetimerEnabled = data.enabled !== false;
          stagetimerVisible = data.visible !== false;
          
          // Update display based on visibility and configuration
          updateStagetimerVisibility();
          
          if (stagetimerEnabled && data.roomId && data.apiKey) {
            startStagetimerPolling();
          } else {
            stopStagetimerPolling();
            updateStagetimerDisplay(null, 'Not configured');
          }
        })
        .catch(err => {
          console.error('Failed to load stagetimer settings:', err);
        });
    }
    
    function updateStagetimerVisibility() {
      const container = document.getElementById('stagetimer-container');
      const hasApiKey = document.getElementById('stagetimer-api-key').value.trim().length > 0;
      const hasRoomId = document.getElementById('stagetimer-room-id').value.trim().length > 0;
      
      // Hide if: not visible OR not enabled OR missing API key/room ID
      if (!stagetimerVisible || !stagetimerEnabled || !hasApiKey || !hasRoomId) {
        container.style.display = 'none';
      } else {
        container.style.display = 'block';
      }
    }
    
    function saveStagetimerSettings() {
      const roomId = document.getElementById('stagetimer-room-id').value.trim();
      const apiKey = document.getElementById('stagetimer-api-key').value.trim();
      const enabled = document.getElementById('stagetimer-enabled').checked;
      const visible = document.getElementById('stagetimer-visible').checked;
      
      fetch(API_BASE + '/api/stagetimer-settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ roomId, apiKey, enabled, visible })
      })
        .then(res => res.json())
        .then(result => {
          if (result.success) {
            showStatus('Stagetimer settings saved', false);
            stagetimerEnabled = enabled;
            stagetimerVisible = visible;
            
            // Update visibility
            updateStagetimerVisibility();
            
            if (stagetimerEnabled && roomId && apiKey) {
              startStagetimerPolling();
            } else {
              stopStagetimerPolling();
              updateStagetimerDisplay(null, enabled ? 'Please configure Room ID and API Key' : 'Disabled');
            }
          } else {
            showStatus('Failed to save: ' + (result.error || 'Unknown error'), true);
          }
        })
        .catch(err => {
          console.error('Save stagetimer settings error:', err);
          showStatus('Failed to save settings: ' + err.message, true);
        });
    }
    
    function updateStagetimerDisplay(data, errorMessage) {
      const container = document.getElementById('stagetimer-container');
      const labelEl = document.getElementById('stagetimer-label');
      const timeEl = document.getElementById('stagetimer-time');
      const statusEl = document.getElementById('stagetimer-status');
      const messagesEl = document.getElementById('stagetimer-messages');
      
      // Check visibility and configuration
      updateStagetimerVisibility();
      
      // If not visible or not enabled, don't update content
      if (!stagetimerVisible || !stagetimerEnabled) {
        return;
      }
      
      // If there's an error and no API key, hide it
      if (errorMessage && errorMessage.includes('not configured')) {
        container.style.display = 'none';
        return;
      }
      
      container.style.display = 'block';
      
      if (errorMessage || !data || !data.success) {
        container.className = 'stagetimer-container error';
        labelEl.textContent = data?.timerName || ''; // Still try to show timer name if available
        timeEl.textContent = '--:--';
        statusEl.textContent = errorMessage || 'Error loading timer';
        messagesEl.style.display = 'none';
        messagesEl.innerHTML = '';
        return;
      }
      
      // Update timer name (label) - use name from API, fallback to empty if not available
      labelEl.textContent = data.timerName || '';
      
      timeEl.textContent = data.displayTime || '0:00';
      
      // Show speaker name below the timer
      statusEl.textContent = data.speaker || '';
      
      // Determine state and styling (but don't show status text)
      if (data.running) {
        container.className = 'stagetimer-container running';
      } else if (data.pause) {
        container.className = 'stagetimer-container';
      } else {
        container.className = 'stagetimer-container';
      }
      
      // Color coding based on remaining time (if available)
      if (data.remainingMs !== undefined) {
        const remainingSeconds = Math.floor(data.remainingMs / 1000);
        if (remainingSeconds <= 15) {
          container.className = 'stagetimer-container critical';
        } else if (remainingSeconds <= 60) {
          container.className = 'stagetimer-container warning';
        }
      }
      
      // Display messages - positioned absolutely so buttons don't move
      console.log('[Stagetimer Display] Updating display, messages:', data.messages);
      if (data.messages && data.messages.length > 0) {
        console.log('[Stagetimer Display] Showing', data.messages.length, 'messages');
        messagesEl.innerHTML = '';
        data.messages.forEach((msg, index) => {
          console.log('[Stagetimer Display] Message', index + ':', msg);
          const messageDiv = document.createElement('div');
          messageDiv.className = 'stagetimer-message ' + (msg.color || 'white');
          if (msg.bold) messageDiv.classList.add('bold');
          if (msg.uppercase) messageDiv.classList.add('uppercase');
          messageDiv.textContent = msg.text || '';
          messagesEl.appendChild(messageDiv);
        });
        messagesEl.classList.add('visible');
      } else {
        console.log('[Stagetimer Display] No messages to display');
        messagesEl.innerHTML = '';
        messagesEl.classList.remove('visible');
      }
    }
    
    function fetchStagetimerStatus() {
      fetch(API_BASE + '/api/get-stagetimer-status')
        .then(res => res.json())
        .then(data => {
          console.log('[Stagetimer] Status response:', JSON.stringify(data, null, 2));
          
          if (!data.configured) {
            updateStagetimerDisplay(null, data.error || 'Not configured');
            return;
          }
          
          if (data.success) {
            console.log('[Stagetimer] Messages in response:', data.messages ? data.messages.length : 0, data.messages);
            
            // Success - reset backoff delay if we were backing off
            if (stagetimerIsBackingOff) {
              stagetimerIsBackingOff = false;
              stagetimerBackoffDelay = 10000; // Reset to 10 seconds
              console.log('[Stagetimer] Rate limit cleared, resetting to 10 second polling');
            }
            
            // Store state for local time calculation
            stagetimerState = {
              success: true,
              running: data.running,
              start: data.start,
              finish: data.finish,
              pause: data.pause,
              serverTime: data.serverTime,
              timerId: data.timerId,
              timerName: data.timerName,
              speaker: data.speaker,
              messages: data.messages || [],
              lastSyncTime: Date.now() // Store when we got this data
            };
            
            // Update display with fresh data (includes messages)
            updateStagetimerDisplay(data, null);
            
            // Also trigger immediate local update to sync display
            updateStagetimerDisplayFromState();
          } else {
            console.error('[Stagetimer] Error:', data.error);
            
            // Handle 429 rate limit errors with exponential backoff
            if (data.error && data.error.includes('429')) {
              if (!stagetimerIsBackingOff) {
                stagetimerIsBackingOff = true;
                // Increase backoff delay exponentially (10s -> 20s -> 40s -> 60s max)
                stagetimerBackoffDelay = Math.min(stagetimerBackoffDelay * 2, 60000);
                console.warn('[Stagetimer] Rate limited (429), backing off to', stagetimerBackoffDelay / 1000, 'seconds');
                // Reschedule with new delay
                scheduleNextStagetimerFetch();
              }
              // Don't clear state on 429 - keep showing last known time
              return;
            }
            
            // For other errors, clear state
            stagetimerState = null;
            updateStagetimerDisplay(null, data.error || 'Error loading timer');
          }
        })
        .catch(err => {
          console.error('Failed to fetch stagetimer status:', err);
          // Don't clear state on network error - keep showing last known time
          // Only update if we have no state at all
          if (!stagetimerState) {
            updateStagetimerDisplay(null, 'Connection error');
          }
        });
    }
    
    function startStagetimerPolling() {
      stopStagetimerPolling();
      stagetimerBackoffDelay = 10000; // Reset backoff on restart
      stagetimerIsBackingOff = false;
      fetchStagetimerStatus(); // Initial fetch
      // Poll API every 10 seconds to sync with server (reduces API calls significantly)
      // This reduces requests to 6/min (much better for rate limits)
      // Will back off further if we get 429 errors
      scheduleNextStagetimerFetch();
      
      // Update display locally every second for smooth counting
      stagetimerDisplayInterval = setInterval(updateStagetimerDisplayFromState, 1000);
    }
    
    function scheduleNextStagetimerFetch() {
      // Clear any existing interval
      if (stagetimerPollInterval) {
        clearTimeout(stagetimerPollInterval);
        stagetimerPollInterval = null;
      }
      
      // Schedule next fetch with current backoff delay
      stagetimerPollInterval = setTimeout(() => {
        fetchStagetimerStatus();
        scheduleNextStagetimerFetch(); // Schedule next one
      }, stagetimerBackoffDelay);
    }
    
    function stopStagetimerPolling() {
      if (stagetimerPollInterval) {
        clearTimeout(stagetimerPollInterval); // Changed from clearInterval to clearTimeout
        stagetimerPollInterval = null;
      }
      if (stagetimerTimerInfoInterval) {
        clearInterval(stagetimerTimerInfoInterval);
        stagetimerTimerInfoInterval = null;
      }
      if (stagetimerDisplayInterval) {
        clearInterval(stagetimerDisplayInterval);
        stagetimerDisplayInterval = null;
      }
      stagetimerBackoffDelay = 10000; // Reset backoff
      stagetimerIsBackingOff = false;
    }
    
    // Calculate and display time locally based on stored state
    function updateStagetimerDisplayFromState() {
      // Check visibility and configuration
      if (!stagetimerVisible || !stagetimerEnabled) {
        return;
      }
      
      if (!stagetimerState || !stagetimerState.success) {
        return; // No state to work with
      }
      
      const state = stagetimerState;
      const now = Date.now();
      
      // Calculate time difference since last server sync
      const timeSinceSync = now - (state.lastSyncTime || now);
      
      // Calculate remaining/elapsed time
      let remainingMs = 0;
      let elapsedMs = 0;
      let displayTime = '0:00';
      let isRunning = state.running || false;
      
      if (state.finish && state.start) {
        const duration = state.finish - state.start;
        
        if (isRunning) {
          // Timer is running - calculate based on server time + elapsed local time
          const serverTimeAtSync = state.serverTime || state.lastSyncTime;
          const localTimeAtSync = state.lastSyncTime;
          const adjustedNow = serverTimeAtSync + (now - localTimeAtSync);
          remainingMs = state.finish - adjustedNow; // Allow negative values
          elapsedMs = adjustedNow - state.start;
        } else if (state.pause) {
          // Timer is paused - use stored values
          elapsedMs = state.pause - state.start;
          remainingMs = duration - elapsedMs;
        } else {
          // Timer not started
          remainingMs = duration;
          elapsedMs = 0;
        }
        
        // Format time as MM:SS or HH:MM:SS (allow negative)
        const totalSeconds = Math.floor(remainingMs / 1000);
        const isNegative = totalSeconds < 0;
        const absSeconds = Math.abs(totalSeconds);
        const hours = Math.floor(absSeconds / 3600);
        const minutes = Math.floor((absSeconds % 3600) / 60);
        const seconds = absSeconds % 60;
        
        const sign = isNegative ? '-' : '';
        const minStr = String(minutes).padStart(2, '0');
        const secStr = String(seconds).padStart(2, '0');
        
        if (hours > 0) {
          displayTime = sign + hours + ':' + minStr + ':' + secStr;
        } else {
          displayTime = sign + minutes + ':' + secStr;
        }
      }
      
      // Update display with calculated time
      const container = document.getElementById('stagetimer-container');
      const labelEl = document.getElementById('stagetimer-label');
      const timeEl = document.getElementById('stagetimer-time');
      const statusEl = document.getElementById('stagetimer-status');
      
      if (!container || !labelEl || !timeEl || !statusEl) return;
      
      labelEl.textContent = state.timerName || '';
      timeEl.textContent = displayTime;
      statusEl.textContent = state.speaker || '';
      
      // Update styling
      if (isRunning) {
        container.className = 'stagetimer-container running';
      } else {
        container.className = 'stagetimer-container';
      }
      
      // Color coding based on remaining time
      const remainingSeconds = Math.floor(remainingMs / 1000);
      if (remainingSeconds <= 15) {
        container.className = 'stagetimer-container critical';
      } else if (remainingSeconds <= 60) {
        container.className = 'stagetimer-container warning';
      }
      
      // Update messages from stored state
      const messagesEl = document.getElementById('stagetimer-messages');
      if (messagesEl && state.messages && state.messages.length > 0) {
        messagesEl.innerHTML = '';
        state.messages.forEach((msg) => {
          const messageDiv = document.createElement('div');
          messageDiv.className = 'stagetimer-message ' + (msg.color || 'white');
          if (msg.bold) messageDiv.classList.add('bold');
          if (msg.uppercase) messageDiv.classList.add('uppercase');
          messageDiv.textContent = msg.text || '';
          messagesEl.appendChild(messageDiv);
        });
        messagesEl.classList.add('visible');
      } else if (messagesEl) {
        messagesEl.innerHTML = '';
        messagesEl.classList.remove('visible');
      }
    }
    
    // Debug console functionality
    const debugConsole = document.getElementById('debug-console');
    const originalConsoleLog = console.log;
    const originalConsoleError = console.error;
    const originalConsoleWarn = console.warn;
    
    function addToDebugConsole(message, type = 'log') {
      const timestamp = new Date().toLocaleTimeString();
      const color = type === 'error' ? '#f44336' : type === 'warn' ? '#ff9800' : '#4caf50';
      const prefix = type === 'error' ? '[ERROR]' : type === 'warn' ? '[WARN]' : '[LOG]';
      
      const logEntry = document.createElement('div');
      logEntry.style.marginBottom = '4px';
      logEntry.style.color = color;
      logEntry.innerHTML = '<span style="color: #888;">[' + timestamp + ']</span> ' + prefix + ' ' + message;
      
      // Remove "Console ready" message if it exists
      const readyMsg = debugConsole.querySelector('div[style*="color: #888"]');
      if (readyMsg && readyMsg.textContent.includes('Console ready')) {
        readyMsg.remove();
      }
      
      debugConsole.appendChild(logEntry);
      debugConsole.scrollTop = debugConsole.scrollHeight;
      
      // Keep only last 100 entries
      while (debugConsole.children.length > 100) {
        debugConsole.removeChild(debugConsole.firstChild);
      }
    }
    
    // Override console methods to also log to debug console
    console.log = function(...args) {
      originalConsoleLog.apply(console, arguments);
      addToDebugConsole(args.map(arg => typeof arg === 'object' ? JSON.stringify(arg, null, 2) : String(arg)).join(' '), 'log');
    };
    
    console.error = function(...args) {
      originalConsoleError.apply(console, arguments);
      addToDebugConsole(args.map(arg => typeof arg === 'object' ? JSON.stringify(arg, null, 2) : String(arg)).join(' '), 'error');
    };
    
    console.warn = function(...args) {
      originalConsoleWarn.apply(console, arguments);
      addToDebugConsole(args.map(arg => typeof arg === 'object' ? JSON.stringify(arg, null, 2) : String(arg)).join(' '), 'warn');
    };
    
    // Clear console button
    document.getElementById('btn-clear-console').addEventListener('click', () => {
      debugConsole.innerHTML = '<div style="color: #888;">Console cleared...</div>';
    });
    
    // Load stagetimer settings on page load
    loadStagetimerSettings();
    
    // Save stagetimer settings button
    document.getElementById('btn-save-stagetimer').addEventListener('click', saveStagetimerSettings);
    document.getElementById('btn-load-stagetimer').addEventListener('click', loadStagetimerSettings);
    
    // Update visibility in real-time when settings change
    document.getElementById('stagetimer-visible').addEventListener('change', () => {
      stagetimerVisible = document.getElementById('stagetimer-visible').checked;
      updateStagetimerVisibility();
    });
    
    document.getElementById('stagetimer-enabled').addEventListener('change', () => {
      stagetimerEnabled = document.getElementById('stagetimer-enabled').checked;
      updateStagetimerVisibility();
    });
    
    document.getElementById('stagetimer-api-key').addEventListener('input', () => {
      updateStagetimerVisibility();
    });
    
    document.getElementById('stagetimer-room-id').addEventListener('input', () => {
      updateStagetimerVisibility();
    });
    
    // Save form
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      
      const data = {
        presentation1: document.getElementById('preset1').value.trim(),
        presentation2: document.getElementById('preset2').value.trim(),
        presentation3: document.getElementById('preset3').value.trim()
      };
      
      fetch(API_BASE + '/api/presets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      })
        .then(res => {
          if (!res.ok) {
            throw new Error('HTTP error! status: ' + res.status);
          }
          return res.json();
        })
        .then(result => {
          if (result.success) {
            showStatus('Presets saved successfully!', false);
            // Reload presets to update the preset buttons
            fetch(API_BASE + '/api/presets')
              .then(res => res.json())
              .then(data => {
                createPresetButtons(data);
              })
              .catch(err => console.error('Failed to reload presets:', err));
          } else {
            showStatus('Failed to save: ' + (result.error || 'Unknown error'), true);
          }
        })
        .catch(err => {
          console.error('Fetch error:', err);
          let errorMsg = 'Failed to save presets: ' + err.message;
          if (err.message.includes('Failed to fetch')) {
            errorMsg += ' (Make sure the app is running and check network connection)';
          }
          showStatus(errorMsg, true);
        });
    });
  </script>
</body>
</html>`;
      
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(html);
      return;
    }
    
    // Proxy API requests to the API server (so Web UI can work over port 80 only)
    if (req.url.startsWith('/api/')) {
      const prefs = loadPreferences();
      const apiPort = prefs.apiPort || DEFAULT_API_PORT;
      
      // Forward the request to the API server
      const apiReq = http.request({
        hostname: '127.0.0.1',
        port: apiPort,
        path: req.url,
        method: req.method,
        headers: {
          'Content-Type': req.headers['content-type'] || 'application/json'
        }
      }, (apiRes) => {
        // Copy response headers
        res.writeHead(apiRes.statusCode, {
          'Content-Type': apiRes.headers['content-type'] || 'application/json',
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type'
        });
        
        // Pipe the response
        apiRes.pipe(res);
      });
      
      apiReq.on('error', (err) => {
        console.error('[Web UI] Proxy error:', err);
        res.writeHead(502, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ 
          success: false, 
          error: 'Cannot connect to API server: ' + err.message 
        }));
      });
      
      // Forward request body if present
      req.pipe(apiReq);
      return;
    }
    
    // 404 for other routes
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not found');
  });
  
  const prefs = loadPreferences();
  const webUiPort = prefs.webUiPort || DEFAULT_WEB_UI_PORT;
  
  webUiServer.listen(webUiPort, '0.0.0.0', () => {
    console.log(`[Web UI] Server listening on http://0.0.0.0:${webUiPort}`);
  }).on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`[Web UI] Port ${webUiPort} is already in use`);
      dialog.showErrorBox(
        'Port Already in Use',
        `Port ${webUiPort} is already in use. Another instance of Google Slides Opener may be running.\n\nPlease quit the other instance or change the Web UI port in settings.`
      );
      // Don't exit the app, but the server won't start
    } else {
      console.error('[Web UI] Server error:', err);
      dialog.showErrorBox(
        'Server Error',
        `Failed to start Web UI server: ${err.message}`
      );
    }
  });
}

app.whenReady().then(() => {
  createWindow();
  startHttpServer();
  startWebUiServer();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('before-quit', () => {
  if (httpServer) {
    console.log('[API] Shutting down HTTP server');
    httpServer.close();
  }
  if (webUiServer) {
    console.log('[Web UI] Shutting down web UI server');
    webUiServer.close();
  }
});
