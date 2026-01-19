const { app, BrowserWindow, ipcMain, screen, session } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow;
let presentationWindow = null;
let notesWindow = null;

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
  const displays = screen.getAllDisplays();
  const presentationDisplay = displays.find(d => d.id === prefs.presentationDisplayId) || displays[0];
  const notesDisplay = displays.find(d => d.id === prefs.notesDisplayId) || displays[0];
  
  console.log('[Test] Using presentation display:', presentationDisplay.id);
  console.log('[Test] Using notes display:', notesDisplay.id);
  
  if (!presentationWindow) {
    presentationWindow = new BrowserWindow({
      x: presentationDisplay.bounds.x,
      y: presentationDisplay.bounds.y,
      width: presentationDisplay.bounds.width,
      height: presentationDisplay.bounds.height,
      fullscreen: true,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        partition: GOOGLE_SESSION_PARTITION
      }
    });

    presentationWindow.on('closed', () => {
      presentationWindow = null;
    });

    // Handle the speaker notes popup window
    presentationWindow.webContents.setWindowOpenHandler((details) => {
      console.log('[Test] Window open intercepted:', details.url);
      console.log('[Test] Frame name:', details.frameName);
      console.log('[Test] Features:', details.features);
      
      // Allow Google Slides to open the speaker notes window
      // Position it on the selected notes display
      const windowOptions = {
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
        
        window.once('ready-to-show', () => {
          console.log('[Test] Window ready-to-show event fired');
          
          if (notesDisplay && notesDisplay.id !== presentationDisplay.id) {
            console.log('[Test] Different displays detected, moving notes window');
            console.log('[Test] Target display bounds:', notesDisplay.bounds);
            
            const targetBounds = {
              x: notesDisplay.bounds.x + 50,
              y: notesDisplay.bounds.y + 50,
              width: notesDisplay.bounds.width - 100,
              height: notesDisplay.bounds.height - 100
            };
            console.log('[Test] Setting window bounds to:', targetBounds);
            
            window.setBounds(targetBounds);
            
            const newBounds = window.getBounds();
            console.log('[Test] Window bounds after setBounds:', newBounds);
            
            setTimeout(() => {
              console.log('[Test] Calling maximize on notes window');
              window.maximize();
              
              const finalBounds = window.getBounds();
              console.log('[Test] Final window bounds after maximize:', finalBounds);
            }, 100);
          } else {
            console.log('[Test] Same display as presentation, just maximizing');
            window.maximize();
          }
        });
        
        app.removeListener('browser-window-created', testWindowCreatedListener);
      }
    };
    app.on('browser-window-created', testWindowCreatedListener);
  }

  presentationWindow.loadURL(testUrl);
  presentationWindow.show();
  
  console.log('[Test] Window opened, loading URL...');
  
  // Listen for all page loads
  presentationWindow.webContents.once('did-finish-load', async () => {
    const currentUrl = presentationWindow.webContents.getURL();
    console.log('[Test] Page loaded:', currentUrl);
    
    console.log('[Test] Waiting 2 seconds before triggering presentation...');
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    console.log('[Test] Attempting to trigger Ctrl+Shift+F5...');
    
    try {
      // Send real keyboard input events
      presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
      presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
      
      console.log('[Test] Ctrl+Shift+F5 sent via sendInputEvent');
    } catch (error) {
      console.error('[Test] Error sending Ctrl+Shift+F5:', error);
    }
    
    // Wait for presentation mode to activate, then press 's' to open speaker notes
    console.log('[Test] Waiting 3 seconds for presentation mode to activate...');
    await new Promise(resolve => setTimeout(resolve, 3000));
    
    console.log('[Test] Attempting to press "s" for speaker notes...');
    
    try {
      // Send real keyboard input events for 's' key
      presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
      presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
      presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
      
      console.log('[Test] "s" key sent via sendInputEvent');
    } catch (error) {
      console.error('[Test] Error sending "s" key:', error);
    }
  });
  
  return { success: true };
});

// Open presentation on specific monitor
ipcMain.handle('open-presentation', async (event, { url, presentationDisplayId, notesDisplayId }) => {
  const displays = screen.getAllDisplays();
  const presentationDisplay = displays.find(d => d.id === presentationDisplayId);
  const notesDisplay = displays.find(d => d.id === notesDisplayId);

  if (!presentationDisplay) {
    return { success: false, message: 'Invalid presentation display' };
  }

  // Close existing windows if any
  if (presentationWindow) presentationWindow.close();
  if (notesWindow) notesWindow.close();

  // Open presentation window
  presentationWindow = new BrowserWindow({
    x: presentationDisplay.bounds.x,
    y: presentationDisplay.bounds.y,
    width: presentationDisplay.bounds.width,
    height: presentationDisplay.bounds.height,
    fullscreen: true,
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
  
  // Listen for new windows being created (this will be the notes window)
  const windowCreatedListener = (event, window) => {
    // Check if this is not the presentation window or main window
    if (window !== presentationWindow && window !== mainWindow) {
      console.log('[Multi-Monitor] Notes window created');
      console.log('[Multi-Monitor] Presentation display ID:', presentationDisplayId);
      console.log('[Multi-Monitor] Notes display ID:', notesDisplayId);
      console.log('[Multi-Monitor] Notes display object:', notesDisplay);
      notesWindow = window;
      
      // Get initial window bounds
      const initialBounds = window.getBounds();
      console.log('[Multi-Monitor] Initial window bounds:', initialBounds);
      
      // Wait for the window to be ready before repositioning
      window.once('ready-to-show', () => {
        console.log('[Multi-Monitor] Window ready-to-show event fired');
        
        // Ensure the window is on the correct display
        if (notesDisplay && notesDisplayId !== presentationDisplayId) {
          console.log('[Multi-Monitor] Different displays detected, moving notes window');
          console.log('[Multi-Monitor] Target display bounds:', notesDisplay.bounds);
          
          const targetBounds = {
            x: notesDisplay.bounds.x + 50,
            y: notesDisplay.bounds.y + 50,
            width: notesDisplay.bounds.width - 100,
            height: notesDisplay.bounds.height - 100
          };
          console.log('[Multi-Monitor] Setting window bounds to:', targetBounds);
          
          // First move the window to the correct display
          window.setBounds(targetBounds);
          
          // Check if it actually moved
          const newBounds = window.getBounds();
          console.log('[Multi-Monitor] Window bounds after setBounds:', newBounds);
          
          // Small delay to ensure the window has moved
          setTimeout(() => {
            console.log('[Multi-Monitor] Calling maximize on notes window');
            window.maximize();
            
            const finalBounds = window.getBounds();
            console.log('[Multi-Monitor] Final window bounds after maximize:', finalBounds);
          }, 100);
        } else {
          console.log('[Multi-Monitor] Same display as presentation, just maximizing');
          window.maximize();
        }
      });
      
      // Remove listener after notes window is created
      app.removeListener('browser-window-created', windowCreatedListener);
    }
  };
  app.on('browser-window-created', windowCreatedListener);

  // Load presentation URL
  presentationWindow.loadURL(url);

  console.log('[Multi-Monitor] Window opened, loading URL...');

  // Listen for all page loads
  presentationWindow.webContents.once('did-finish-load', async () => {
    const currentUrl = presentationWindow.webContents.getURL();
    console.log('[Multi-Monitor] Page loaded:', currentUrl);
    
    console.log('[Multi-Monitor] Waiting 2 seconds before triggering presentation...');
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    console.log('[Multi-Monitor] Attempting to trigger Ctrl+Shift+F5...');
    
    try {
      // Send real keyboard input events
      presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F5', modifiers: ['control', 'shift'] });
      presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'F5', modifiers: ['control', 'shift'] });
      
      console.log('[Multi-Monitor] Ctrl+Shift+F5 sent via sendInputEvent');
    } catch (error) {
      console.error('[Multi-Monitor] Error sending Ctrl+Shift+F5:', error);
    }
    
    // Wait for presentation mode to activate, then press 's' to open speaker notes
    console.log('[Multi-Monitor] Waiting 3 seconds for presentation mode to activate...');
    await new Promise(resolve => setTimeout(resolve, 3000));
    
    console.log('[Multi-Monitor] Attempting to press "s" for speaker notes...');
    
    try {
      // Send real keyboard input events for 's' key
      presentationWindow.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S' });
      presentationWindow.webContents.sendInputEvent({ type: 'char', keyCode: 's' });
      presentationWindow.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S' });
      
      console.log('[Multi-Monitor] "s" key sent via sendInputEvent');
    } catch (error) {
      console.error('[Multi-Monitor] Error sending "s" key:', error);
    }
  });

  presentationWindow.on('closed', () => {
    presentationWindow = null;
  });

  return { success: true };
});

app.whenReady().then(() => {
  createWindow();

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
