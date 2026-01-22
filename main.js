const { app, BrowserWindow, ipcMain, screen, session } = require('electron');
const path = require('path');
const fs = require('fs');
const http = require('http');
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

// Minimize the left preview pane in speaker notes window by moving the divider to the left
function minimizeSpeakerNotesPreviewPane(window) {
  if (!window || window.isDestroyed()) return;
  
  // Wait for page to load, then try to resize the divider
  window.webContents.once('did-finish-load', () => {
    // Try multiple times with increasing delays to catch the DOM when it's ready
    [500, 1000, 1500].forEach((delay, index) => {
      setTimeout(() => {
        if (window.isDestroyed()) return;
        
        window.webContents.executeJavaScript(`
          (function() {
            try {
              // Strategy 1: Find and drag the divider/resizer element
              var divider = null;
              var allElements = document.querySelectorAll('*');
              
              // Look for elements with resize cursor or draggable attribute
              for (var i = 0; i < allElements.length; i++) {
                var el = allElements[i];
                var style = window.getComputedStyle(el);
                var rect = el.getBoundingClientRect();
                
                // Check if element looks like a vertical divider (thin, vertical, has resize cursor)
                if ((style.cursor === 'ew-resize' || style.cursor === 'col-resize' || 
                     el.draggable === true || el.getAttribute('role') === 'separator') &&
                    rect.width < 20 && rect.height > 100) {
                  divider = el;
                  break;
                }
              }
              
              if (divider) {
                // Found divider - simulate dragging it to the left
                var rect = divider.getBoundingClientRect();
                var startX = rect.left + rect.width / 2;
                var targetX = 150; // Target position for left edge of right pane
                
                // Create mouse events to simulate drag
                var mousedown = new MouseEvent('mousedown', {
                  bubbles: true,
                  cancelable: true,
                  view: window,
                  clientX: startX,
                  clientY: rect.top + rect.height / 2,
                  button: 0,
                  buttons: 1
                });
                
                var mousemove = new MouseEvent('mousemove', {
                  bubbles: true,
                  cancelable: true,
                  view: window,
                  clientX: targetX,
                  clientY: rect.top + rect.height / 2,
                  button: 0,
                  buttons: 1
                });
                
                var mouseup = new MouseEvent('mouseup', {
                  bubbles: true,
                  cancelable: true,
                  view: window,
                  clientX: targetX,
                  clientY: rect.top + rect.height / 2,
                  button: 0,
                  buttons: 0
                });
                
                divider.dispatchEvent(mousedown);
                setTimeout(function() {
                  divider.dispatchEvent(mousemove);
                  setTimeout(function() {
                    divider.dispatchEvent(mouseup);
                  }, 100);
                }, 50);
                
                return { success: true, method: 'divider-drag', attempt: ${index + 1} };
              }
              
              // Strategy 2: Find left pane and set its width directly
              var leftPane = null;
              var panes = document.querySelectorAll('div[style*="width"], div[style*="flex"]');
              
              for (var i = 0; i < panes.length; i++) {
                var pane = panes[i];
                var rect = pane.getBoundingClientRect();
                var style = window.getComputedStyle(pane);
                
                // Look for a pane that's on the left side and contains slide previews
                if (rect.left < window.innerWidth * 0.4 && 
                    (pane.textContent.includes('Slide') || pane.querySelector('img') || 
                     pane.querySelector('[class*="preview"]') || pane.querySelector('[class*="slide"]'))) {
                  leftPane = pane;
                  break;
                }
              }
              
              if (leftPane) {
                // Set left pane to minimum width
                leftPane.style.width = '150px';
                leftPane.style.minWidth = '150px';
                leftPane.style.maxWidth = '200px';
                leftPane.style.flexBasis = '150px';
                leftPane.style.flexShrink = '0';
                
                // Also try to find and adjust any flex container
                var container = leftPane.parentElement;
                if (container) {
                  var containerStyle = window.getComputedStyle(container);
                  if (containerStyle.display === 'flex') {
                    // Force the left pane to be small
                    leftPane.style.flex = '0 0 150px';
                  }
                }
                
                return { success: true, method: 'direct-pane-resize', attempt: ${index + 1} };
              }
              
              // Strategy 3: Look for CSS variables or data attributes that control width
              var containers = document.querySelectorAll('[style*="--"], [data-width], [style*="grid"]');
              for (var i = 0; i < containers.length; i++) {
                var container = containers[i];
                var style = container.style;
                
                // Try to set CSS custom properties if they exist
                if (style.getPropertyValue('--left-pane-width')) {
                  style.setProperty('--left-pane-width', '150px');
                  return { success: true, method: 'css-variable', attempt: ${index + 1} };
                }
              }
              
              return { success: false, error: 'Could not find divider or left pane', attempt: ${index + 1} };
            } catch (e) {
              return { success: false, error: e.message, attempt: ${index + 1} };
            }
          })()
        `).then(result => {
          if (result && result.success) {
            console.log('[Notes] Successfully minimized preview pane:', result.method, '(attempt', result.attempt + ')');
          } else if (index === 2) {
            // Only log failure on last attempt to avoid spam
            console.log('[Notes] Could not minimize preview pane:', result ? result.error : 'unknown error');
          }
        }).catch(err => {
          if (index === 2) {
            console.log('[Notes] Error minimizing preview pane:', err.message);
          }
        });
      }, delay);
    });
  });
}


// Use a persistent session for Google authentication
const GOOGLE_SESSION_PARTITION = 'persist:google';

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
      // Position it on the selected notes display
      const windowOptions = {
        frame: false,
        webPreferences: {
          nodeIntegration: false,
          contextIsolation: true,
          partition: GOOGLE_SESSION_PARTITION
        }
      };
      
      if (notesDisplay && notesDisplay.id !== presentationDisplay.id) {
        windowOptions.x = notesDisplay.bounds.x;
        windowOptions.y = notesDisplay.bounds.y;
        windowOptions.width = notesDisplay.bounds.width;
        windowOptions.height = notesDisplay.bounds.height;
      } else {
        windowOptions.width = 1280;
        windowOptions.height = 720;
      }
      
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
        
        window.once('ready-to-show', () => {
          console.log('[Test] Window ready-to-show event fired');
          
          // Minimize the left preview pane in speaker notes
          minimizeSpeakerNotesPreviewPane(window);
          
          // Always move the window to the correct display first, then maximize
          if (notesDisplay) {
            const targetBounds = {
              x: notesDisplay.bounds.x + 50,
              y: notesDisplay.bounds.y + 50,
              width: notesDisplay.bounds.width - 100,
              height: notesDisplay.bounds.height - 100
            };
            
            if (notesDisplay.id !== presentationDisplay.id) {
              console.log('[Test] Different displays detected, moving notes window to display:', notesDisplay.id);
            } else {
              console.log('[Test] Same display as presentation, but still moving to ensure correct position');
            }
            
            console.log('[Test] Target display bounds:', notesDisplay.bounds);
            console.log('[Test] Setting window bounds to:', targetBounds);
            
            window.setBounds(targetBounds);
            
            const newBounds = window.getBounds();
            console.log('[Test] Window bounds after setBounds:', newBounds);
            
            setTimeout(() => {
              console.log('[Test] Calling maximize on notes window');
              window.maximize();
              
              const finalBounds = window.getBounds();
              console.log('[Test] Final window bounds after maximize:', finalBounds);
              
              // Log actual final positions after a short delay
              setTimeout(() => {
                const actualPresentationBounds = presentationWindow ? presentationWindow.getBounds() : null;
                const actualNotesBounds = notesWindow ? notesWindow.getBounds() : null;
                console.log('[Test] ===== ACTUAL FINAL WINDOW POSITIONS =====');
                console.log('[Test] Presentation window actual position:', actualPresentationBounds);
                console.log('[Test] Notes window actual position:', actualNotesBounds);
                console.log('[Test] =============================================');
              }, 500);
            }, 50);
          }
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
  
  // Set up navigation listener to detect presentation mode activation
  let sKeyPressed = false;
  const navigationListener = async (event, url) => {
    console.log('[Test] Navigated to:', url);
    
    // Check if we're in presentation mode (URL contains /present/ or /localpresent but not /presentation/)
    const isPresentMode = (url.includes('/present/') || url.includes('/localpresent')) && !url.includes('/presentation/');
    if (isPresentMode && !sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
      sKeyPressed = true;
      console.log('[Test] Presentation mode URL detected, pressing "s" for speaker notes...');
      
      // Small delay to ensure presentation mode UI is ready
      await new Promise(resolve => setTimeout(resolve, 300));
      
      if (presentationWindow && !presentationWindow.isDestroyed()) {
        try {
          // Focus the window to ensure it receives the keyboard events
          presentationWindow.focus();
          await new Promise(resolve => setTimeout(resolve, 50));
          
          // Send real keyboard input events for 's' key
          presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
          presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
          presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
          
          console.log('[Test] "s" key sent via sendInputEvent');
        } catch (error) {
          console.error('[Test] Error sending "s" key:', error);
        }
        
        // Remove the listener after we've pressed 's'
        presentationWindow.webContents.removeListener('did-navigate', navigationListener);
      }
    }
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
    
    // Fallback: if navigation doesn't detect presentation mode, press 's' after a delay
    setTimeout(async () => {
      if (!sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
        console.log('[Test] Fallback timer: pressing "s" for speaker notes...');
        sKeyPressed = true;
        
        try {
          presentationWindow.focus();
          await new Promise(resolve => setTimeout(resolve, 50));
          
          presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
          presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
          presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
          
          console.log('[Test] "s" key sent via sendInputEvent (fallback)');
        } catch (error) {
          console.error('[Test] Error sending "s" key (fallback):', error);
        }
        
        if (presentationWindow && !presentationWindow.isDestroyed()) {
          presentationWindow.webContents.removeListener('did-navigate', navigationListener);
        }
      }
    }, 1000);
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
    // Position it on the selected notes display
    const windowOptions = {
      frame: false,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        partition: GOOGLE_SESSION_PARTITION
      }
    };
    
    if (notesDisplay && notesDisplayIdNum !== presentationDisplayIdNum) {
      windowOptions.x = notesDisplay.bounds.x;
      windowOptions.y = notesDisplay.bounds.y;
      windowOptions.width = notesDisplay.bounds.width;
      windowOptions.height = notesDisplay.bounds.height;
    } else {
      windowOptions.width = 1280;
      windowOptions.height = 720;
    }
    
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
      
      // Wait for the window to be ready before repositioning
      window.once('ready-to-show', () => {
        console.log('[Multi-Monitor] Window ready-to-show event fired');
        
        // Minimize the left preview pane in speaker notes
        minimizeSpeakerNotesPreviewPane(window);
        
        // Always move the window to the correct display first, then maximize
        if (notesDisplay) {
          const targetBounds = {
            x: notesDisplay.bounds.x + 50,
            y: notesDisplay.bounds.y + 50,
            width: notesDisplay.bounds.width - 100,
            height: notesDisplay.bounds.height - 100
          };
          
          if (notesDisplayIdNum !== presentationDisplayIdNum) {
            console.log('[Multi-Monitor] Different displays detected, moving notes window to display:', notesDisplayIdNum);
          } else {
            console.log('[Multi-Monitor] Same display as presentation, but still moving to ensure correct position');
          }
          
          console.log('[Multi-Monitor] Target display bounds:', notesDisplay.bounds);
          console.log('[Multi-Monitor] Setting window bounds to:', targetBounds);
          
          window.setBounds(targetBounds);
          
          const newBounds = window.getBounds();
          console.log('[Multi-Monitor] Window bounds after setBounds:', newBounds);
          
          setTimeout(() => {
            console.log('[Multi-Monitor] Calling maximize on notes window');
            window.maximize();
            
            const finalBounds = window.getBounds();
            console.log('[Multi-Monitor] Final window bounds after maximize:', finalBounds);
            
            // Log actual final positions after a short delay
            setTimeout(() => {
              const actualPresentationBounds = presentationWindow ? presentationWindow.getBounds() : null;
              const actualNotesBounds = notesWindow ? notesWindow.getBounds() : null;
              console.log('[Multi-Monitor] ===== ACTUAL FINAL WINDOW POSITIONS =====');
              console.log('[Multi-Monitor] Presentation window actual position:', actualPresentationBounds);
              console.log('[Multi-Monitor] Notes window actual position:', actualNotesBounds);
              console.log('[Multi-Monitor] =============================================');
            }, 500);
          }, 50);
        }
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
  let sKeyPressed = false;
  const navigationListener = async (event, url) => {
    console.log('[Multi-Monitor] Navigated to:', url);
    
    // Check if we're in presentation mode (URL contains /present/ or /localpresent but not /presentation/)
    const isPresentMode = (url.includes('/present/') || url.includes('/localpresent')) && !url.includes('/presentation/');
    if (isPresentMode && !sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
      sKeyPressed = true;
      console.log('[Multi-Monitor] Presentation mode URL detected, pressing "s" for speaker notes...');
      
      // Small delay to ensure presentation mode UI is ready
      await new Promise(resolve => setTimeout(resolve, 300));
      
      if (presentationWindow && !presentationWindow.isDestroyed()) {
        try {
          // Focus the window to ensure it receives the keyboard events
          presentationWindow.focus();
          await new Promise(resolve => setTimeout(resolve, 50));
          
          // Send real keyboard input events for 's' key
          presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
          presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
          presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
          
          console.log('[Multi-Monitor] "s" key sent via sendInputEvent');
        } catch (error) {
          console.error('[Multi-Monitor] Error sending "s" key:', error);
        }
        
        // Remove the listener after we've pressed 's'
        presentationWindow.webContents.removeListener('did-navigate', navigationListener);
      }
    }
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
    
    // Fallback: if navigation doesn't detect presentation mode, press 's' after a delay
    setTimeout(async () => {
      if (!sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
        console.log('[Multi-Monitor] Fallback timer: pressing "s" for speaker notes...');
        sKeyPressed = true;
        
        try {
          presentationWindow.focus();
          await new Promise(resolve => setTimeout(resolve, 50));
          
          presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
          presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
          presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
          
          console.log('[Multi-Monitor] "s" key sent via sendInputEvent (fallback)');
        } catch (error) {
          console.error('[Multi-Monitor] Error sending "s" key (fallback):', error);
        }
        
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
          version: '1.2.5',
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
            
            const windowOptions = {
              frame: false,
              webPreferences: {
                nodeIntegration: false,
                contextIsolation: true,
                partition: GOOGLE_SESSION_PARTITION
              }
            };
            
            if (notesDisplay && notesDisplayId !== presentationDisplayId) {
              windowOptions.x = notesDisplay.bounds.x;
              windowOptions.y = notesDisplay.bounds.y;
              windowOptions.width = notesDisplay.bounds.width;
              windowOptions.height = notesDisplay.bounds.height;
            } else {
              windowOptions.width = 1280;
              windowOptions.height = 720;
            }
            
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
              
              // Position and maximize notes window
              window.once('ready-to-show', () => {
                // Minimize the left preview pane in speaker notes
                minimizeSpeakerNotesPreviewPane(window);
                
                if (notesDisplay) {
                  const targetBounds = {
                    x: notesDisplay.bounds.x + 50,
                    y: notesDisplay.bounds.y + 50,
                    width: notesDisplay.bounds.width - 100,
                    height: notesDisplay.bounds.height - 100
                  };
                  
                  window.setBounds(targetBounds);
                  
                  setTimeout(() => {
                    window.maximize();
                  }, 50);
                }
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
            
            // Fallback timer
            setTimeout(async () => {
              console.log('[API] Fallback timer triggered, sKeyPressed:', sKeyPressed);
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
          // Capture current slide
          const savedSlide = typeof currentSlide === 'number' ? currentSlide : 1;
          const urlToReload = lastPresentationUrl;

          // Capture current zoom level from notes window before closing
          let savedZoomLevel = null;
          
          // Function to restore zoom level (will be called after notes window opens)
          const restoreZoomLevel = async (window) => {
            if (savedZoomLevel === null || !window || window.isDestroyed()) return;
            
            console.log('[API] Restoring zoom level to:', savedZoomLevel);
            
            // Wait for the page to fully load
            await new Promise(resolve => setTimeout(resolve, 1500));
            
            try {
              // Get current zoom level
              const currentZoom = await window.webContents.executeJavaScript(`
                (function() {
                  try {
                    var zoomInBtn = document.querySelector('[title="Zoom in"]');
                    var zoomOutBtn = document.querySelector('[title="Zoom out"]');
                    var contentArea = document.querySelector('[role="main"]') || document.body;
                    var style = window.getComputedStyle(contentArea);
                    var fontSize = parseFloat(style.fontSize);
                    var transform = style.transform;
                    var scale = 1;
                    if (transform && transform !== 'none') {
                      var match = transform.match(/scale\\(([^)]+)\\)/);
                      if (match) scale = parseFloat(match[1]);
                    }
                    var allText = document.body.innerText || '';
                    var zoomMatch = allText.match(/(\\d+)%/);
                    var zoomPercent = zoomMatch ? parseInt(zoomMatch[1], 10) : null;
                    
                    return {
                      fontSize: fontSize,
                      scale: scale,
                      zoomPercent: zoomPercent,
                      currentZoom: zoomPercent || (scale * 100) || Math.round((fontSize / 14) * 100)
                    };
                  } catch (e) {
                    return { currentZoom: 100 };
                  }
                })()
              `);
              
              const targetZoom = savedZoomLevel;
              const currentZoomValue = currentZoom.currentZoom || 100;
              const zoomDifference = targetZoom - currentZoomValue;
              
              console.log('[API] Current zoom:', currentZoomValue, 'Target zoom:', targetZoom, 'Difference:', zoomDifference);
              
              // Calculate how many clicks needed (each click is roughly 10-20% change)
              // We'll use a more conservative estimate of ~15% per click
              const clicksNeeded = Math.round(zoomDifference / 15);
              
              if (Math.abs(clicksNeeded) > 0 && Math.abs(zoomDifference) > 5) {
                const buttonToClick = clicksNeeded > 0 ? 'Zoom in' : 'Zoom out';
                const clickCount = Math.abs(clicksNeeded);
                
                console.log('[API] Clicking', buttonToClick, clickCount, 'times');
                
                for (let i = 0; i < clickCount; i++) {
                  await new Promise(resolve => setTimeout(resolve, 250)); // Delay between clicks
                  
                  const result = await window.webContents.executeJavaScript(`
                    (function() {
                      try {
                        var btn = document.querySelector('[title="${buttonToClick}"]');
                        if (btn && !btn.disabled && btn.getAttribute('aria-disabled') !== 'true') {
                          var mousedownEvent = new MouseEvent('mousedown', {
                            bubbles: true,
                            cancelable: true,
                            view: window,
                            button: 0
                          });
                          var mouseupEvent = new MouseEvent('mouseup', {
                            bubbles: true,
                            cancelable: true,
                            view: window,
                            button: 0
                          });
                          var clickEvent = new MouseEvent('click', {
                            bubbles: true,
                            cancelable: true,
                            view: window,
                            button: 0
                          });
                          
                          btn.dispatchEvent(mousedownEvent);
                          btn.dispatchEvent(mouseupEvent);
                          btn.dispatchEvent(clickEvent);
                          
                          return { success: true };
                        }
                        return { success: false, error: 'Button not found or disabled' };
                      } catch (e) {
                        return { success: false, error: e.message };
                      }
                    })()
                  `);
                  
                  if (!result.success) {
                    console.log('[API] Zoom button click failed:', result.error);
                    break; // Stop if button becomes unavailable
                  }
                }
                
                console.log('[API] Zoom restoration complete');
              } else {
                console.log('[API] Zoom level already correct (difference:', zoomDifference, '), no adjustment needed');
              }
            } catch (error) {
              console.error('[API] Error restoring zoom level:', error);
            }
          };
          if (notesWindow && !notesWindow.isDestroyed()) {
            try {
              const zoomInfo = await notesWindow.webContents.executeJavaScript(`
                (function() {
                  try {
                    // Try to detect zoom level by checking font sizes or transform scales
                    // Look for the main content area in speaker notes
                    var contentArea = document.querySelector('[role="main"]') || 
                                    document.querySelector('.speaker-notes') ||
                                    document.body;
                    
                    if (contentArea) {
                      var style = window.getComputedStyle(contentArea);
                      var fontSize = parseFloat(style.fontSize);
                      var transform = style.transform;
                      
                      // Check for zoom/scale in transform
                      var scale = 1;
                      if (transform && transform !== 'none') {
                        var match = transform.match(/scale\\(([^)]+)\\)/);
                        if (match) {
                          scale = parseFloat(match[1]);
                        }
                      }
                      
                      // Try to find zoom buttons to see if they're disabled (at min/max)
                      var zoomInBtn = document.querySelector('[title="Zoom in"]');
                      var zoomOutBtn = document.querySelector('[title="Zoom out"]');
                      var zoomInDisabled = zoomInBtn ? zoomInBtn.disabled || zoomInBtn.getAttribute('aria-disabled') === 'true' : false;
                      var zoomOutDisabled = zoomOutBtn ? zoomOutBtn.disabled || zoomOutBtn.getAttribute('aria-disabled') === 'true' : false;
                      
                      // Try to find a zoom indicator in the UI
                      var zoomText = null;
                      var allText = document.body.innerText || '';
                      var zoomMatch = allText.match(/(\\d+)%/);
                      if (zoomMatch) {
                        zoomText = parseInt(zoomMatch[1], 10);
                      }
                      
                      return {
                        fontSize: fontSize,
                        scale: scale,
                        zoomInDisabled: zoomInDisabled,
                        zoomOutDisabled: zoomOutDisabled,
                        zoomPercent: zoomText,
                        // Calculate relative zoom: we'll use a combination of factors
                        relativeZoom: zoomText || (scale * 100) || null
                      };
                    }
                    return null;
                  } catch (e) {
                    return null;
                  }
                })()
              `);
              
              if (zoomInfo && zoomInfo.relativeZoom !== null) {
                savedZoomLevel = zoomInfo.relativeZoom;
                console.log('[API] Captured zoom level:', savedZoomLevel);
              } else {
                // Fallback: try to count zoom clicks by checking button states
                // If zoom out is disabled, we're at minimum; if zoom in is disabled, we're at maximum
                // Otherwise, we'll try to detect from font size
                if (zoomInfo) {
                  if (zoomInfo.zoomOutDisabled && !zoomInfo.zoomInDisabled) {
                    savedZoomLevel = 50; // Minimum zoom
                  } else if (zoomInfo.zoomInDisabled && !zoomInfo.zoomOutDisabled) {
                    savedZoomLevel = 200; // Maximum zoom
                  } else if (zoomInfo.fontSize) {
                    // Use font size as a proxy (default is usually around 14-16px)
                    savedZoomLevel = Math.round((zoomInfo.fontSize / 14) * 100);
                  } else {
                    savedZoomLevel = 100; // Default/unknown
                  }
                  console.log('[API] Captured zoom level (fallback):', savedZoomLevel);
                }
              }
            } catch (error) {
              console.error('[API] Error capturing zoom level:', error);
              savedZoomLevel = 100; // Default fallback
            }
          }

          console.log('[API] Reloading presentation: saving slide', savedSlide, 'zoom level', savedZoomLevel, 'URL:', urlToReload);

          // Close existing windows
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

          // Wait a moment for windows to close
          await new Promise(resolve => setTimeout(resolve, 500));

          // Re-open the presentation (reuse the same logic as /api/open-presentation)
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
            if (notesDisplay && notesDisplayId !== presentationDisplayId) {
              windowOptions.x = notesDisplay.bounds.x;
              windowOptions.y = notesDisplay.bounds.y;
              windowOptions.width = notesDisplay.bounds.width;
              windowOptions.height = notesDisplay.bounds.height;
            } else {
              windowOptions.width = 1280;
              windowOptions.height = 720;
            }
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
              window.once('ready-to-show', () => {
                // Minimize the left preview pane in speaker notes
                minimizeSpeakerNotesPreviewPane(window);
                
                if (notesDisplay) {
                  const targetBounds = {
                    x: notesDisplay.bounds.x + 50,
                    y: notesDisplay.bounds.y + 50,
                    width: notesDisplay.bounds.width - 100,
                    height: notesDisplay.bounds.height - 100
                  };
                  window.setBounds(targetBounds);
                  setTimeout(() => { window.maximize(); }, 50);
                }
                
                // Restore zoom level if we saved one
                if (savedZoomLevel !== null) {
                  // Wait for the page to fully load before restoring zoom
                  window.webContents.once('did-finish-load', () => {
                    restoreZoomLevel(window);
                  });
                }
              });
              app.removeListener('browser-window-created', windowCreatedListener);
            }
          };
          app.on('browser-window-created', windowCreatedListener);

          // Set up navigation listener for speaker notes
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

          // Load the presentation
          const presentUrl = toPresentUrl(urlToReload);
          lastPresentationUrl = urlToReload; // Re-store it
          currentSlide = 1; // Will be updated after navigation
          
          // Ensure fullscreen on macOS (sometimes needed after programmatic window creation)
          // Set up fullscreen handler before loading
          presentationWindow.once('ready-to-show', () => {
            if (process.platform === 'darwin' && presentationWindow && !presentationWindow.isDestroyed()) {
              // Ensure window is on correct display and set fullscreen
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
          
          presentationWindow.loadURL(presentUrl);
          presentationWindow.show();
          
          // Fallback: set fullscreen after a short delay if ready-to-show already fired
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

          // Wait for presentation to load and enter presentation mode
          await new Promise(resolve => setTimeout(resolve, 2000));

          // Trigger presentation mode (Ctrl+Shift+F5)
          presentationWindow.webContents.once('did-finish-load', async () => {
            await new Promise(resolve => setTimeout(resolve, 200));
            if (presentationWindow && !presentationWindow.isDestroyed()) {
              presentationWindow.focus();
              await new Promise(resolve => setTimeout(resolve, 50));
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
            }
          });

          // Fallback: press 's' for speaker notes after delay
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

          // Wait for presentation mode to be ready, then navigate to saved slide
          await new Promise(resolve => setTimeout(resolve, 3000));

          if (presentationWindow && !presentationWindow.isDestroyed() && savedSlide > 1) {
            console.log('[API] Navigating to saved slide:', savedSlide);
            presentationWindow.focus();
            await new Promise(resolve => setTimeout(resolve, 100));

            // Navigate to the saved slide (from slide 1)
            const slidesToMove = savedSlide - 1;
            for (let i = 0; i < slidesToMove; i++) {
              presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Right' });
              presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Right' });
              if (i < slidesToMove - 1) {
                await new Promise(resolve => setTimeout(resolve, 150));
              }
            }
            currentSlide = savedSlide;
            console.log('[API] Reload complete: returned to slide', savedSlide);
          }

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

    // POST /api/open-speaker-notes - Toggle speaker notes (s key)
    if (req.method === 'POST' && req.url === '/api/open-speaker-notes') {
      if (!presentationWindow || presentationWindow.isDestroyed()) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'No presentation is open' }));
        return;
      }

      try {
        presentationWindow.focus();
        presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
        presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
        presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, message: 'Speaker notes toggled' }));
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
            if (notesDisplay && notesDisplayId !== presentationDisplayId) {
              windowOptions.x = notesDisplay.bounds.x;
              windowOptions.y = notesDisplay.bounds.y;
              windowOptions.width = notesDisplay.bounds.width;
              windowOptions.height = notesDisplay.bounds.height;
            } else {
              windowOptions.width = 1280;
              windowOptions.height = 720;
            }
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
              window.once('ready-to-show', () => {
                minimizeSpeakerNotesPreviewPane(window);
                if (notesDisplay) {
                  const targetBounds = {
                    x: notesDisplay.bounds.x + 50,
                    y: notesDisplay.bounds.y + 50,
                    width: notesDisplay.bounds.width - 100,
                    height: notesDisplay.bounds.height - 100
                  };
                  window.setBounds(targetBounds);
                  setTimeout(() => { window.maximize(); }, 50);
                }
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
    .container {
      background: white;
      border-radius: 16px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3);
      max-width: 600px;
      width: 100%;
      padding: 40px;
    }
    h1 {
      color: #333;
      margin-bottom: 10px;
      font-size: 28px;
    }
    .subtitle {
      color: #666;
      margin-bottom: 30px;
      font-size: 14px;
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
      margin-top: 20px;
      padding: 12px;
      border-radius: 8px;
      text-align: center;
      font-size: 14px;
      display: none;
    }
    .status.success {
      background: #d4edda;
      color: #155724;
      border: 1px solid #c3e6cb;
      display: block;
    }
    .status.error {
      background: #f8d7da;
      color: #721c24;
      border: 1px solid #f5c6cb;
      display: block;
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
      transition: all 0.2s;
      margin-bottom: -2px;
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
  </style>
</head>
<body>
  <div class="container">
    <h1>${machineName}</h1>
    <p class="subtitle">Control your presentations</p>
    
    <!-- Tabs -->
    <div class="tabs">
      <button class="tab-btn active" data-tab="controls">Controls</button>
      <button class="tab-btn" data-tab="settings">Settings</button>
    </div>
    
    <!-- Controls Tab (Default) -->
    <div id="tab-controls" class="tab-content active">
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
      
      <!-- Slide Controls -->
      <div class="controls-section">
        <h3>Slide Controls</h3>
        <div class="controls-grid">
          <button type="button" class="btn-control" id="btn-prev-slide" title="Go to previous slide">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="15 18 9 12 15 6"></polyline>
            </svg>
            Previous Slide
          </button>
          <button type="button" class="btn-control" id="btn-next-slide" title="Go to next slide">
            Next Slide
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="9 18 15 12 9 6"></polyline>
            </svg>
          </button>
          <button type="button" class="btn-control" id="btn-reload" title="Reload presentation and return to current slide">
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="23 4 23 10 17 10"></polyline>
              <polyline points="1 20 1 14 7 14"></polyline>
              <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path>
            </svg>
            Reload Presentation
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
    </div>
    
    <div id="status" class="status"></div>
  </div>
  
  <script>
    const form = document.getElementById('preset-form');
    const loadBtn = document.getElementById('load-btn');
    const status = document.getElementById('status');
    // Use current hostname so it works from other machines, fallback to localhost
    const API_BASE = 'http://' + (window.location.hostname || '127.0.0.1') + ':' + ${apiPort};
    
    function showStatus(message, isError) {
      status.textContent = message;
      status.className = 'status ' + (isError ? 'error' : 'success');
      setTimeout(() => {
        status.className = 'status';
      }, 3000);
    }
    
    // Prevent native tooltips and use custom floating ones
    document.querySelectorAll('.btn-control[title]').forEach(btn => {
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
      return fetch(API_BASE + endpoint, {
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
          if (result.success !== false) {
            showStatus(result.message || 'Action completed successfully', false);
          } else {
            showStatus(result.error || 'Action failed', true);
          }
          return result;
        })
        .catch(err => {
          console.error('API call error:', err);
          showStatus('Failed: ' + err.message + ' (Make sure the app is running)', true);
          throw err;
        });
    }
    
    // Set up control buttons
    document.getElementById('btn-next-slide').addEventListener('click', () => {
      apiCall('/api/next-slide');
    });
    
    document.getElementById('btn-prev-slide').addEventListener('click', () => {
      apiCall('/api/previous-slide');
    });
    
    document.getElementById('btn-reload').addEventListener('click', () => {
      apiCall('/api/reload-presentation');
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
    
    // Load current presets on page load
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
      })
      .catch(err => {
        console.error('Failed to load presets:', err);
        // Don't show error on initial load, just log it
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
        })
        .catch(err => {
          console.error('Load error:', err);
          showStatus('Failed to load presets: ' + err.message + ' (Make sure the app is running)', true);
        });
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
    
    // 404 for other routes
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not found');
  });
  
  const prefs = loadPreferences();
  const webUiPort = prefs.webUiPort || DEFAULT_WEB_UI_PORT;
  
  webUiServer.listen(webUiPort, '0.0.0.0', () => {
    console.log(`[Web UI] Server listening on http://0.0.0.0:${webUiPort}`);
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
