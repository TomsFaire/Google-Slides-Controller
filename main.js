const { app, BrowserWindow, ipcMain, screen, session } = require('electron');
const path = require('path');
const fs = require('fs');
const http = require('http');

let mainWindow;
let presentationWindow = null;
let notesWindow = null;
let currentSlide = null; // best-effort: we track on our next/prev; DOM can override when notes window has aria-posinset/aria-setsize

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
    if (fs.existsSync(prefsPath)) {
      const data = fs.readFileSync(prefsPath, 'utf8');
      return JSON.parse(data);
    }
  } catch (error) {
    console.error('Error loading preferences:', error);
  }
  return {};
}

// Save preferences
function savePreferences(prefs) {
  try {
    const prefsPath = getPreferencesPath();
    fs.writeFileSync(prefsPath, JSON.stringify(prefs, null, 2), 'utf8');
  } catch (error) {
    console.error('Error saving preferences:', error);
  }
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 600,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    },
    autoHideMenuBar: true,
    resizable: false,
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

// Save preferences
ipcMain.handle('save-preferences', async (event, prefs) => {
  savePreferences(prefs);
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
  const googleSession = getGoogleSession();
  const cookies = await googleSession.cookies.get({ domain: '.google.com' });
  
  // Check if we have Google authentication cookies
  const hasAuthCookies = cookies.some(cookie => 
    cookie.name === 'SID' || cookie.name === 'HSID' || cookie.name === 'SSID'
  );
  
  return { signedIn: hasAuthCookies };
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
const API_PORT = 9595;
let httpServer;

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
        const state = {
          status: 'ok',
          version: '1.0.0',
          presentationOpen: !!(presentationWindow && !presentationWindow.isDestroyed()),
          notesOpen: !!(notesWindow && !notesWindow.isDestroyed()),
          currentSlide: currentSlide,
          totalSlides: null
        };
        if (notesWindow && !notesWindow.isDestroyed()) {
          try {
            const info = await notesWindow.webContents.executeJavaScript(`
              (function(){
                var el = document.querySelector('[aria-posinset]');
                if (!el) return null;
                var cur = parseInt(el.getAttribute('aria-posinset'), 10);
                var tot = parseInt(el.getAttribute('aria-setsize'), 10);
                return { current: isNaN(cur) ? null : cur, total: isNaN(tot) ? null : tot };
              })()
            `);
            if (info && info.current != null) {
              state.currentSlide = info.current;
              if (info.total != null) state.totalSlides = info.total;
            }
          } catch (e) { /* DOM not available or changed */ }
        }
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
          
          // Set up navigation listener to detect presentation mode
          let sKeyPressed = false;
          const navigationListener = async (event, navUrl) => {
            console.log('[API] Navigated to:', navUrl);
            const isPresentMode = (navUrl.includes('/present/') || navUrl.includes('localpresent')) && !navUrl.includes('/presentation/');
            if (isPresentMode && !sKeyPressed && presentationWindow && !presentationWindow.isDestroyed()) {
              console.log('[API] Presentation mode detected, pressing "s" for speaker notes');
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
          currentSlide = 1;
          presentationWindow.loadURL(presentUrl);
          presentationWindow.show();
          // Send response immediately
          if (!res.headersSent) {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: true, message: 'Presentation opened' }));
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
    
    // 404 for unknown endpoints
    res.writeHead(404, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Not found' }));
  });
  
  httpServer.listen(API_PORT, '0.0.0.0', () => {
    console.log(`[API] HTTP server listening on http://0.0.0.0:${API_PORT}`);
  });
}

app.whenReady().then(() => {
  createWindow();
  startHttpServer();

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
});
