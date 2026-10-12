// EDEN BI — Windows desktop app (Electron). A window around the live
// dashboard: data and updates come from https://dashboard.edenalkhalij.com,
// so the app never needs re-installing when the dashboard changes.
const { app, BrowserWindow, shell, Menu, dialog } = require('electron');
const path = require('path');

const HOME = 'https://dashboard.edenalkhalij.com/';
const SAME = u => { try { const h = new URL(u).hostname; return h === 'dashboard.edenalkhalij.com'; } catch (e) { return false; } };
let win = null;

if (!app.requestSingleInstanceLock()) { app.quit(); }
app.on('second-instance', () => { if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });
app.setAppUserModelId('com.edenalkhalij.dashboard');

function popupWindow(parent) {
  // print / PDF previews opened by the dashboard (window.open)
  return {
    action: 'allow',
    overrideBrowserWindowOptions: {
      parent, width: 1100, height: 800, autoHideMenuBar: true, title: 'EDEN BI',
      icon: path.join(__dirname, 'build', 'icon.png'),
      webPreferences: { contextIsolation: true, nodeIntegration: false, partition: 'persist:eden' }
    }
  };
}

function wire(contents) {
  contents.setWindowOpenHandler(({ url }) => {
    if (!url || url === 'about:blank' || url.startsWith('blob:') || url.startsWith('data:') || SAME(url)) return popupWindow(win);
    if (/^(https?|mailto|tel):/i.test(url)) shell.openExternal(url);   // WhatsApp, maps, sheets… in their own apps
    return { action: 'deny' };
  });
  contents.on('will-navigate', (e, url) => {
    if (SAME(url) || url.startsWith('file:')) return;
    if (/^https:\/\/(accounts\.google\.com|script\.google\.com|script\.googleusercontent\.com)/i.test(url)) return;
    e.preventDefault(); shell.openExternal(url);
  });
}

function createWindow() {
  win = new BrowserWindow({
    width: 1440, height: 900, minWidth: 900, minHeight: 600, show: false,
    title: 'EDEN BI', backgroundColor: '#f4f0e8', autoHideMenuBar: true,
    icon: path.join(__dirname, 'build', 'icon.png'),
    webPreferences: { contextIsolation: true, nodeIntegration: false, spellcheck: false, partition: 'persist:eden' }
  });
  win.maximize();
  win.once('ready-to-show', () => win.show());
  wire(win.webContents);
  win.webContents.on('did-fail-load', (e, code, desc, url, isMain) => {
    if (isMain && code !== -3) win.loadFile(path.join(__dirname, 'offline.html'));   // -3 = aborted, ignore
  });
  win.loadURL(HOME);
}

app.on('web-contents-created', (e, contents) => { if (contents.getType() === 'window') wire(contents); });

function handleDownloads() {
  const { session } = require('electron');
  session.fromPartition('persist:eden').on('will-download', (ev, item) => {
    // Excel / CSV / PDF exports go straight to Downloads and the folder opens.
    const fs = require('fs'), dir = app.getPath('downloads'), name = item.getFilename();
    const ext = path.extname(name), base = path.basename(name, ext);
    let dest = path.join(dir, name), n = 1;
    while (fs.existsSync(dest)) dest = path.join(dir, `${base} (${n++})${ext}`);
    item.setSavePath(dest);
    item.once('done', (_e, state) => { if (state === 'completed') shell.showItemInFolder(dest); });
  });
}

const isMac = process.platform === 'darwin';
Menu.setApplicationMenu(Menu.buildFromTemplate([
  { label: 'EDEN BI', submenu: [
    { label: 'Home', accelerator: 'Alt+Home', click: () => win && win.loadURL(HOME) },
    { role: 'reload', label: 'Refresh' }, { role: 'forceReload', label: 'Hard refresh' },
    { type: 'separator' }, { role: 'togglefullscreen' }, { role: 'zoomIn' }, { role: 'zoomOut' }, { role: 'resetZoom' },
    { type: 'separator' }, { label: 'About', click: () => dialog.showMessageBox(win, { title: 'EDEN BI', message: 'EDEN Business Intelligence', detail: 'Version ' + app.getVersion() + '\n' + HOME }) },
    isMac ? { role: 'close' } : { role: 'quit', label: 'Exit' }
  ] },
  { label: 'Edit', submenu: [{ role: 'undo' }, { role: 'redo' }, { type: 'separator' }, { role: 'cut' }, { role: 'copy' }, { role: 'paste' }, { role: 'selectAll' }] }
]));

app.whenReady().then(() => { handleDownloads(); createWindow(); });
app.on('window-all-closed', () => app.quit());
