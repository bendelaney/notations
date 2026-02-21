const { app, BrowserWindow, ipcMain, dialog, nativeTheme, Menu, clipboard, nativeImage } = require('electron');
const fsNative = require('fs');
const fs = require('fs/promises');
const path = require('path');

const DEV_WATCH = process.env.NOTATIONS_DEV_WATCH === '1';
const PROCESS_RESTART_FILES = new Set(['main.js', 'preload.js']);
const STATE_FILENAME = 'notations-state.json';
const IS_MAC = process.platform === 'darwin';
const DEEP_LINK_PROTOCOL = 'notations';
const APP_NAME = 'Notations';
const ICONS_DIR = path.join(__dirname, 'assets', 'icons');
const APP_ICON_PATH = path.join(ICONS_DIR, 'icon.png');

app.setName(APP_NAME);
process.title = APP_NAME;

let pendingDeepLinkPath = null;
let mainWindow = null;

function parseDeepLinkUrl(urlString) {
  if (!urlString) return null;
  try {
    const parsed = new URL(urlString);
    if (parsed.protocol !== `${DEEP_LINK_PROTOCOL}:`) {
      return null;
    }
    const segments = [];
    if (parsed.hostname) {
      segments.push(parsed.hostname);
    }
    const pathSegments = parsed.pathname.split('/').filter(Boolean);
    segments.push(...pathSegments);
    return `/${segments.map((s) => decodeURIComponent(s)).join('/')}`;
  } catch (_) {
    return null;
  }
}

function findDeepLinkInArgv(argv = []) {
  for (let i = 0; i < argv.length; i += 1) {
    const value = String(argv[i] || '');
    if (value.startsWith(`${DEEP_LINK_PROTOCOL}://`)) {
      return parseDeepLinkUrl(value);
    }
  }
  return null;
}

function convertWindowUrlToDeepLink(urlString) {
  try {
    const parsed = new URL(urlString);
    let route = String(parsed.hash || '').replace(/^#/, '');
    if (!route || route === '/') {
      route = parsed.searchParams.get('route') || '/';
    }
    const segments = route
      .split('/')
      .filter(Boolean)
      .map((s) => encodeURIComponent(decodeURIComponent(s)));
    return segments.length ? `${DEEP_LINK_PROTOCOL}://${segments.join('/')}` : `${DEEP_LINK_PROTOCOL}://`;
  } catch (_) {
    return `${DEEP_LINK_PROTOCOL}://`;
  }
}

function dispatchDeepLink(routePath) {
  if (!routePath) return;
  const targetWindow = mainWindow && !mainWindow.isDestroyed() ? mainWindow : BrowserWindow.getAllWindows()[0];
  if (!targetWindow || targetWindow.isDestroyed()) {
    pendingDeepLinkPath = routePath;
    return;
  }
  if (targetWindow.webContents.isLoading()) {
    pendingDeepLinkPath = routePath;
    return;
  }
  targetWindow.webContents.send('open-deep-link', routePath);
}

function registerDeepLinkProtocol() {
  if (process.defaultApp && process.argv.length >= 2) {
    app.setAsDefaultProtocolClient(DEEP_LINK_PROTOCOL, process.execPath, [path.resolve(process.argv[1])]);
    return;
  }
  app.setAsDefaultProtocolClient(DEEP_LINK_PROTOCOL);
}

function getStateFilePath() {
  return path.join(app.getPath('userData'), STATE_FILENAME);
}

function applyAppIcon() {
  if (!IS_MAC || !app.dock || !fsNative.existsSync(APP_ICON_PATH)) return;
  const icon = nativeImage.createFromPath(APP_ICON_PATH);
  if (!icon.isEmpty()) {
    app.dock.setIcon(icon);
  }
}

function applyAboutPanelOptions() {
  if (!IS_MAC) return;
  const options = {
    applicationName: APP_NAME,
    applicationVersion: app.getVersion(),
    version: app.getVersion()
  };
  if (fsNative.existsSync(APP_ICON_PATH)) {
    options.iconPath = APP_ICON_PATH;
  }
  app.setAboutPanelOptions(options);
}

function setupRendererWatch(win) {
  if (!DEV_WATCH) return () => {};

  const targets = [
    path.join(__dirname, 'app', 'renderer'),
    path.join(__dirname, 'fonts.css'),
    path.join(__dirname, 'assets', 'icons')
  ];

  const watchers = [];
  let reloadTimer = null;

  const queueReload = () => {
    if (reloadTimer) {
      clearTimeout(reloadTimer);
    }
    reloadTimer = setTimeout(() => {
      if (win.isDestroyed()) return;
      win.webContents.reloadIgnoringCache();
    }, 70);
  };

  targets.forEach((target) => {
    if (!fsNative.existsSync(target)) return;
    try {
      const watcher = fsNative.watch(target, { recursive: true }, queueReload);
      watchers.push(watcher);
    } catch (_) {
      // ignore watch setup failures for non-critical paths
    }
  });

  return () => {
    if (reloadTimer) {
      clearTimeout(reloadTimer);
    }
    watchers.forEach((watcher) => watcher.close());
  };
}

function setupProcessRelaunchWatch() {
  if (!DEV_WATCH) return () => {};

  let relaunchTimer = null;
  let relaunching = false;
  let watcher = null;

  const queueRelaunch = () => {
    if (relaunching) return;
    if (relaunchTimer) {
      clearTimeout(relaunchTimer);
    }
    relaunchTimer = setTimeout(() => {
      if (relaunching) return;
      relaunching = true;
      app.relaunch();
      app.exit(0);
    }, 120);
  };

  try {
    watcher = fsNative.watch(__dirname, (eventType, filename) => {
      if (!filename) return;
      const changed = String(filename);
      if (!PROCESS_RESTART_FILES.has(changed)) return;
      queueRelaunch();
    });
  } catch (_) {
    // ignore process-watch setup failures in non-dev environments
  }

  return () => {
    if (relaunchTimer) {
      clearTimeout(relaunchTimer);
    }
    if (watcher) {
      watcher.close();
    }
  };
}

function createWindow() {
  nativeTheme.themeSource = 'light';

  const win = new BrowserWindow({
    width: 1440,
    height: 960,
    backgroundColor: '#ffffff',
    icon: APP_ICON_PATH,
    show: false,
    title: APP_NAME,
    titleBarStyle: IS_MAC ? 'hidden' : 'default',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  mainWindow = win;
  win.once('ready-to-show', () => win.show());
  win.loadFile(path.join(__dirname, 'app', 'renderer', 'index.html'));
  win.webContents.on('did-finish-load', () => {
    if (pendingDeepLinkPath) {
      const routePath = pendingDeepLinkPath;
      pendingDeepLinkPath = null;
      win.webContents.send('open-deep-link', routePath);
    }
  });

  const stopWatching = setupRendererWatch(win);
  win.on('closed', () => {
    stopWatching();
    if (mainWindow === win) {
      mainWindow = null;
    }
  });
}

function copyFocusedWindowLink() {
  const win = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
  if (!win || win.isDestroyed()) {
    return;
  }
  const url = win.webContents.getURL();
  if (url) {
    clipboard.writeText(convertWindowUrlToDeepLink(url));
  }
}

function requestFocusedWindowPrint() {
  const win = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
  if (!win || win.isDestroyed()) {
    return;
  }
  win.webContents.send('request-print-document');
}

function createApplicationMenu() {
  const template = [
    ...(IS_MAC
      ? [
          {
            label: APP_NAME,
            submenu: [
              {
                label: `About ${APP_NAME}`,
                click: () => app.showAboutPanel()
              },
              { type: 'separator' },
              { role: 'services' },
              { type: 'separator' },
              { role: 'hide' },
              { role: 'hideOthers' },
              { role: 'unhide' },
              { type: 'separator' },
              { role: 'quit' }
            ]
          }
        ]
      : []),
    {
      label: 'File',
      submenu: [
        {
          label: 'Print...',
          accelerator: 'CmdOrCtrl+P',
          click: requestFocusedWindowPrint
        },
        { type: 'separator' },
        {
          label: 'Copy Link',
          accelerator: 'CmdOrCtrl+Shift+L',
          click: copyFocusedWindowLink
        },
        { type: 'separator' },
        { role: IS_MAC ? 'close' : 'quit' }
      ]
    },
    {
      label: 'Edit',
      submenu: [
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'selectAll' }
      ]
    },
    {
      label: 'View',
      submenu: [{ role: 'reload' }, { role: 'forceReload' }, { role: 'toggleDevTools' }, { type: 'separator' }, { role: 'resetZoom' }, { role: 'zoomIn' }, { role: 'zoomOut' }, { type: 'separator' }, { role: 'togglefullscreen' }]
    },
    {
      label: 'Window',
      submenu: [{ role: 'minimize' }, { role: 'zoom' }, ...(IS_MAC ? [{ type: 'separator' }, { role: 'front' }] : [{ role: 'close' }])]
    }
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

const initialDeepLink = findDeepLinkInArgv(process.argv);
if (initialDeepLink) {
  pendingDeepLinkPath = initialDeepLink;
}

const gotSingleInstanceLock = app.requestSingleInstanceLock();
if (!gotSingleInstanceLock) {
  app.quit();
}

if (gotSingleInstanceLock) {
  app.on('second-instance', (_event, argv) => {
    const routePath = findDeepLinkInArgv(argv);
    if (routePath) {
      dispatchDeepLink(routePath);
    }
    const targetWindow = mainWindow && !mainWindow.isDestroyed() ? mainWindow : BrowserWindow.getAllWindows()[0];
    if (targetWindow) {
      if (targetWindow.isMinimized()) {
        targetWindow.restore();
      }
      targetWindow.focus();
    }
  });
}

app.on('open-url', (event, urlString) => {
  event.preventDefault();
  const routePath = parseDeepLinkUrl(urlString);
  if (routePath) {
    dispatchDeepLink(routePath);
  }
});

app.whenReady().then(() => {
  const stopProcessWatch = setupProcessRelaunchWatch();
  app.on('will-quit', stopProcessWatch);
  applyAboutPanelOptions();
  createApplicationMenu();
  registerDeepLinkProtocol();
  applyAppIcon();

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

ipcMain.handle('print-document', async (event, payload = {}) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (!win) {
    return { ok: false, reason: 'Window unavailable' };
  }

  const safePayload = payload && typeof payload === 'object' ? payload : {};
  const rawMargins = safePayload.margins && typeof safePayload.margins === 'object' ? safePayload.margins : null;
  const hasCustomMargins =
    !!rawMargins && ['top', 'right', 'bottom', 'left'].every((key) => Number.isFinite(Number(rawMargins[key])));

  const printOptions = {
    silent: false,
    printBackground: true,
    landscape: false,
    scaleFactor: 100,
    margins: { marginType: 'default' }
  };

  if (typeof safePayload.pageSize === 'string' && safePayload.pageSize.trim()) {
    printOptions.pageSize = safePayload.pageSize.trim();
  }

  if (hasCustomMargins) {
    printOptions.margins = {
      marginType: 'custom',
      top: Math.max(0, Math.round(Number(rawMargins.top))),
      right: Math.max(0, Math.round(Number(rawMargins.right))),
      bottom: Math.max(0, Math.round(Number(rawMargins.bottom))),
      left: Math.max(0, Math.round(Number(rawMargins.left)))
    };
  }

  const warning = await dialog.showMessageBox(win, {
    type: 'warning',
    buttons: ['Print at 100%', 'Cancel'],
    defaultId: 0,
    cancelId: 1,
    title: 'Print Scale',
    message: 'Use 100% print scale only. Do not use "Fit to page".',
    detail: 'Notations relies on exact sizing so the printed page matches what you see on screen.'
  });

  if (warning.response !== 0) {
    return { ok: false, cancelled: true };
  }

  const printWithOptions = (options) =>
    new Promise((resolve) => {
      win.webContents.print(options, (success, failureReason) => {
        if (!success) {
          resolve({ ok: false, reason: failureReason || 'Print failed' });
          return;
        }
        resolve({ ok: true });
      });
    });

  let result = await printWithOptions(printOptions);
  if (!result.ok && hasCustomMargins) {
    const fallbackOptions = {
      ...printOptions,
      margins: { marginType: 'default' }
    };
    result = await printWithOptions(fallbackOptions);
  }

  return result;
});

ipcMain.handle('export-pdf', async (event, payload = {}) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (!win) {
    return { ok: false, reason: 'Window unavailable' };
  }

  const defaultStem = payload.defaultFilename || 'notation';

  const save = await dialog.showSaveDialog(win, {
    title: 'Export PDF',
    defaultPath: `${defaultStem}.pdf`,
    filters: [{ name: 'PDF', extensions: ['pdf'] }]
  });

  if (save.canceled || !save.filePath) {
    return { ok: false, cancelled: true };
  }

  try {
    const data = await win.webContents.printToPDF({
      printBackground: true,
      landscape: false,
      preferCSSPageSize: true
    });

    await fs.writeFile(save.filePath, data);
    return { ok: true, filePath: save.filePath };
  } catch (error) {
    return { ok: false, reason: error.message };
  }
});

ipcMain.handle('export-text', async (event, payload = {}) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (!win) {
    return { ok: false, reason: 'Window unavailable' };
  }

  const body = typeof payload.body === 'string' ? payload.body : '';
  const defaultStem = payload.defaultFilename || 'notation';

  const save = await dialog.showSaveDialog(win, {
    title: 'Export Text',
    defaultPath: `${defaultStem}.txt`,
    filters: [{ name: 'Text', extensions: ['txt'] }]
  });

  if (save.canceled || !save.filePath) {
    return { ok: false, cancelled: true };
  }

  try {
    await fs.writeFile(save.filePath, body, 'utf8');
    return { ok: true, filePath: save.filePath };
  } catch (error) {
    return { ok: false, reason: error.message };
  }
});

ipcMain.handle('load-state', async () => {
  try {
    const statePath = getStateFilePath();
    const raw = await fs.readFile(statePath, 'utf8');
    return { ok: true, state: JSON.parse(raw) };
  } catch (error) {
    if (error && error.code === 'ENOENT') {
      return { ok: true, state: null };
    }
    return { ok: false, reason: error.message };
  }
});

ipcMain.handle('save-state', async (_event, payload = {}) => {
  try {
    const serialized =
      typeof payload.serialized === 'string'
        ? payload.serialized
        : JSON.stringify(payload.state || {});
    const statePath = getStateFilePath();
    await fs.mkdir(path.dirname(statePath), { recursive: true });
    await fs.writeFile(statePath, serialized, 'utf8');
    return { ok: true };
  } catch (error) {
    return { ok: false, reason: error.message };
  }
});

ipcMain.handle('consume-initial-deep-link', async () => {
  const routePath = pendingDeepLinkPath || null;
  pendingDeepLinkPath = null;
  return { ok: true, routePath };
});
