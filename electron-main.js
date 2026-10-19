'use strict';
/*
 * Electron main process for Savvy Repair for Microsoft Office.
 *
 * Loads the self-contained web app from web/index.html using file://
 * protocol. All repair logic runs in the renderer with no Node.js access.
 * External links are forwarded to the system browser.
 */
const { app, BrowserWindow, shell } = require('electron');
const path = require('path');

app.on('window-all-closed', () => {
  // On macOS apps conventionally stay active until the user quits via Cmd+Q.
  if (process.platform !== 'darwin') app.quit();
});

function createWindow() {
  const win = new BrowserWindow({
    width: 960,
    height: 740,
    minWidth: 600,
    minHeight: 480,
    title: 'Savvy Repair for Microsoft Office',
    icon: path.join(__dirname, 'web', 'icons', 'icon-512.png'),
    webPreferences: {
      nodeIntegration: false,     // renderer has no Node.js access
      contextIsolation: true,
      sandbox: true,
      // Service workers require https:// or localhost://, not file://.
      // Disable registration so the console is not littered with SW errors.
      serviceWorkers: false,
    },
  });

  // Open any http/https link (GitHub, etc.) in the default system browser,
  // not inside the Electron window.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) {
      shell.openExternal(url);
      return { action: 'deny' };
    }
    return { action: 'allow' };
  });

  win.webContents.on('will-navigate', (event, url) => {
    if (/^https?:/.test(url)) {
      event.preventDefault();
      shell.openExternal(url);
    }
  });

  win.loadFile(path.join(__dirname, 'web', 'index.html'));
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    // On macOS re-create the window when the dock icon is clicked and no
    // windows are open.
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});
