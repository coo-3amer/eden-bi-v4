// EDEN BI — Windows desktop app (Electron).
// The window shows the live dashboard (https://dashboard.edenalkhalij.com), so
// dashboard changes appear without reinstalling. Around it the app adds:
// splash screen, EDEN title bar, remembered window size/place, tray icon,
// Windows notifications, taskbar badge, start with Windows, keyboard
// shortcuts, self-update, and the last data when there is no internet.
const { app, BrowserWindow, WebContentsView, shell, ipcMain, Tray, Menu, Notification,
  nativeImage, screen, session, net, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const HOME = process.env.EDEN_BI_URL || 'https://dashboard.edenalkhalij.com/';
const HOST = new URL(HOME).host;
const PROTO = new URL(HOME).protocol.replace(':', '');
const BAR_H = 44;
const PART = 'persist:eden';
const ICON = path.join(__dirname, 'build', 'icon.png');
const HIDDEN = process.argv.includes('--hidden');

let win = null, dash = null, splash = null, tray = null, quitting = false, shown = false;
let state = { crumb: '', theme: 'light', badge: 0, loggedIn: false, offline: false, lastOnline: 0,
  maximized: false, canBack: false, canFwd: false, update: '' };

if (!app.requestSingleInstanceLock()) { app.quit(); process.exit(0); }
app.on('second-instance', () => showMain());
app.setAppUserModelId('com.edenalkhalij.dashboard');

/* ---------- small JSON store in the user's app folder ---------- */
const DATA = () => app.getPath('userData');
function readJson(name, def) { try { return JSON.parse(fs.readFileSync(path.join(DATA(), name), 'utf8')); } catch (e) { return def; } }
function writeJson(name, v) { try { fs.writeFileSync(path.join(DATA(), name), JSON.stringify(v)); } catch (e) {} }
let prefs = null;
const pref = (k, d) => (prefs || (prefs = readJson('prefs.json', {})))[k] ?? d;
const setPref = (k, v) => { pref(k); prefs[k] = v; writeJson('prefs.json', prefs); };

/* ---------- 3 · window size and place ---------- */
function savedBounds() {
  const b = readJson('window.json', null);
  if (!b || !b.width) return null;
  const ok = screen.getAllDisplays().some(d => {
    const a = d.workArea;
    return b.x < a.x + a.width - 80 && b.x + b.width > a.x + 80 && b.y >= a.y - 10 && b.y < a.y + a.height - 60;
  });
  return ok ? b : null;
}
let saveTimer = null;
function saveBounds() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    if (!win || win.isDestroyed() || win.isMinimized()) return;
    const max = win.isMaximized(), b = max ? (readJson('window.json', {}) || {}) : win.getNormalBounds();
    writeJson('window.json', Object.assign({}, b, { maximized: max }));
  }, 400);
}

/* ---------- 1 · splash ---------- */
function createSplash() {
  splash = new BrowserWindow({ width: 620, height: 380, frame: false, transparent: true, resizable: false,
    movable: true, show: false, skipTaskbar: false, alwaysOnTop: false, center: true, icon: ICON, title: 'EDEN BI',
    webPreferences: { contextIsolation: true, nodeIntegration: false } });
  splash.loadFile(path.join(__dirname, 'ui', 'splash.html'), { query: { v: app.getVersion() } });
  splash.once('ready-to-show', () => splash && splash.show());
}
function splashSay(text, stuck) {
  if (splash && !splash.isDestroyed()) splash.webContents.executeJavaScript(`window.say&&say(${JSON.stringify(text)},${!!stuck})`).catch(() => {});
}
function closeSplash() { if (splash && !splash.isDestroyed()) splash.close(); splash = null; }

/* ---------- 2 · main window: EDEN title bar + dashboard below it ---------- */
function layout() {
  if (!win || !dash) return;
  const [w, h] = win.getContentSize();
  dash.setBounds({ x: 0, y: BAR_H, width: w, height: Math.max(0, h - BAR_H) });
}
function createMain() {
  const b = savedBounds();
  win = new BrowserWindow({
    width: b ? b.width : 1440, height: b ? b.height : 900, x: b ? b.x : undefined, y: b ? b.y : undefined,
    minWidth: 900, minHeight: 600, frame: false, show: false, title: 'EDEN BI', icon: ICON,
    backgroundColor: '#f4f0e8', thickFrame: true,
    webPreferences: { preload: path.join(__dirname, 'ui', 'titlebar-preload.js'), contextIsolation: true, nodeIntegration: false }
  });
  win.loadFile(path.join(__dirname, 'ui', 'titlebar.html'));
  dash = new WebContentsView({ webPreferences: { preload: path.join(__dirname, 'ui', 'dash-preload.js'),
    contextIsolation: true, nodeIntegration: false, spellcheck: false, partition: PART } });
  dash.setBackgroundColor('#f4f0e8');
  win.contentView.addChildView(dash);
  layout();
  win.on('resize', () => { layout(); saveBounds(); });
  win.on('move', saveBounds);
  win.on('maximize', () => { push({ maximized: true }); saveBounds(); });
  win.on('unmaximize', () => { push({ maximized: false }); saveBounds(); });
  win.on('close', e => {                       // 4 · closing keeps EDEN BI running in the tray
    if (quitting) return;
    e.preventDefault(); win.hide();
    if (!pref('trayHintShown', false)) {
      setPref('trayHintShown', true);
      notify('EDEN BI is still running', 'Alerts keep coming. Right-click the EDEN icon near the clock to quit.');
    }
  });
  if (b && b.maximized) win.maximize();
  else if (!b) win.maximize();

  wire(dash.webContents);
  dash.webContents.on('before-input-event', (e, input) => shortcut(e, input));
  win.webContents.on('before-input-event', (e, input) => shortcut(e, input));
  dash.webContents.on('did-finish-load', () => {
    readyToShow();
  });
  dash.webContents.on('did-fail-load', (e, code, desc, url, isMain) => {
    if (!isMain || code === -3) return;
    if (!shown) { splashSay('No connection — retrying…', true); setTimeout(() => dash.webContents.loadURL(HOME), 5000); }
    else dash.webContents.loadFile(path.join(__dirname, 'offline.html'));
  });
  dash.webContents.loadURL(HOME);
}
function readyToShow() {
  if (shown) return;
  shown = true;
  setTimeout(() => {
    if (!HIDDEN) { win.show(); win.focus(); }
    closeSplash();
  }, 900);
}
function showMain(go) {
  if (!win) return;
  if (win.isMinimized()) win.restore();
  win.show(); win.focus();
  if (go) openPage(go);
}

/* ---------- links, pop-ups (print / PDF), downloads ---------- */
const same = u => { try { return new URL(u).host === HOST; } catch (e) { return false; } };
function wire(contents) {
  contents.setWindowOpenHandler(({ url }) => {
    if (!url || url === 'about:blank' || url.startsWith('blob:') || url.startsWith('data:') || same(url)) {
      return { action: 'allow', overrideBrowserWindowOptions: { width: 1100, height: 800, autoHideMenuBar: true,
        title: 'EDEN BI', icon: ICON, webPreferences: { contextIsolation: true, nodeIntegration: false, partition: PART } } };
    }
    if (/^(https?|mailto|tel):/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  contents.on('will-navigate', (e, url) => {
    if (same(url) || url.startsWith('file:')) return;
    if (/^https:\/\/(accounts\.google\.com|script\.google\.com|script\.googleusercontent\.com)/i.test(url)) return;
    e.preventDefault(); shell.openExternal(url);
  });
}
function handleDownloads(ses) {
  ses.on('will-download', (ev, item) => {
    const dir = app.getPath('downloads'), name = item.getFilename(), ext = path.extname(name), base = path.basename(name, ext);
    let dest = path.join(dir, name), n = 1;
    while (fs.existsSync(dest)) dest = path.join(dir, `${base} (${n++})${ext}`);
    item.setSavePath(dest);
    item.once('done', (_e, st) => { if (st === 'completed') shell.showItemInFolder(dest); });
  });
}

/* ---------- 10 · offline: keep the last copy of the page and its data ---------- */
const CACHE = () => { const d = path.join(DATA(), 'offline'); fs.mkdirSync(d, { recursive: true }); return d; };
const keyFile = k => path.join(CACHE(), crypto.createHash('sha1').update(k).digest('hex') + '.bin');
const STATIC_HOSTS = /^(cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|unpkg\.com|fonts\.googleapis\.com|fonts\.gstatic\.com)$/i;
function offlineCache(ses) {
  ses.protocol.handle(PROTO, async req => {
    const u = new URL(req.url);
    const cacheable = req.method === 'GET' && (u.host === HOST || STATIC_HOSTS.test(u.hostname));
    if (!cacheable) return net.fetch(req, { bypassCustomProtocolHandlers: true }).catch(() => Response.error());
    const key = u.host === HOST ? u.origin + u.pathname : req.url;
    try {
      // the dashboard's own files are always re-checked with the server (cheap 304 when
      // unchanged), so a new version shows on the next refresh instead of up to 10 min later
      const res = u.host === HOST
        ? await net.fetch(req.url, { headers: req.headers, cache: 'no-cache', bypassCustomProtocolHandlers: true })
        : await net.fetch(req, { bypassCustomProtocolHandlers: true });
      if (res.ok) {
        const buf = Buffer.from(await res.clone().arrayBuffer());
        const meta = { type: res.headers.get('content-type') || '', at: Date.now() };
        fs.writeFile(keyFile(key), buf, () => {});
        fs.writeFile(keyFile(key) + '.json', JSON.stringify(meta), () => {});
      }
      return res;
    } catch (err) {
      try {
        const buf = fs.readFileSync(keyFile(key));
        const meta = JSON.parse(fs.readFileSync(keyFile(key) + '.json', 'utf8'));
        if (u.host === HOST) setOffline(true);
        return new Response(buf, { status: 200, headers: { 'content-type': meta.type || 'application/octet-stream' } });
      } catch (e) { throw err; }
    }
  });
}
function setOffline(off) {
  if (!off) state.lastOnline = Date.now(), writeJson('online.json', { at: state.lastOnline });
  if (state.offline !== off) push({ offline: off });
}
ipcMain.handle('cache:get', (e, key) => { try { return fs.readFileSync(keyFile('api|' + key), 'utf8'); } catch (x) { return null; } });
ipcMain.on('cache:set', (e, key, text) => { fs.writeFile(keyFile('api|' + key), String(text), () => {}); });
ipcMain.on('net:state', (e, off) => setOffline(!!off));

/* ---------- reading the dashboard (page name, theme, badge, login) ---------- */
const PAGE_PROBE = `(() => {
  const t = document.documentElement.dataset.theme || 'light';
  const logged = typeof AUTH_TOKEN !== 'undefined' && !!AUTH_TOKEN;
  const ov = {bc:'Broker Commissions',sc:'Sales Scorecard',da:'Discount Approvals',pm:'Price Manager',al:'Activity Log',ew:'EDEN WALK Partners'};
  let crumb = '';
  const mod = el => { const m = el && el.closest('.module'); const h = m && m.querySelector('.moduleHeader span'); return h ? h.textContent.trim() : ''; };
  const o = Object.keys(ov).find(k => document.getElementById(k));
  if (o) { const b = document.querySelector('.moduleMenu [data-' + o + ']'); crumb = (mod(b) ? mod(b) + ' › ' : '') + ov[o]; }
  else { const a = document.querySelector('.moduleMenu button.active'); if (a) crumb = (mod(a) ? mod(a) + ' › ' : '') + a.textContent.replace(/\\d+$/,'').trim(); }
  if (!logged) crumb = 'Sign in';
  const c = document.querySelector('.moduleMenu [data-da] .daCount');
  return { theme: t, loggedIn: logged, crumb, badge: c ? (parseInt(c.textContent, 10) || 0) : 0 };
})()`;
const history = [];
let histPos = -1, histJump = false;
async function probe() {
  if (!dash || dash.webContents.isDestroyed() || dash.webContents.isLoading()) return;
  try {
    const s = await dash.webContents.executeJavaScript(PAGE_PROBE, true);
    if (s.crumb && s.crumb !== state.crumb && s.loggedIn) {
      if (histJump) histJump = false;
      else { history.splice(histPos + 1); history.push(s.crumb); if (history.length > 50) history.shift(); histPos = history.length - 1; }
    }
    if (s.badge !== state.badge) setBadge(s.badge);
    push(Object.assign(s, { canBack: histPos > 0, canFwd: histPos < history.length - 1 }));
  } catch (e) {}
}
function push(patch) {
  Object.assign(state, patch || {});
  if (win && !win.isDestroyed()) win.webContents.send('state', state);
  if (tray) tray.setToolTip('EDEN BI' + (state.badge ? ` — ${state.badge} discount approval${state.badge > 1 ? 's' : ''} waiting` : '') + (state.offline ? ' — offline' : ''));
}

/* open a page by its menu name (sidebar button text or pop-up page) */
function openPage(name) {
  if (!dash) return;
  const js = `(() => {
    const want = ${JSON.stringify(String(name).split('›').pop().trim().toLowerCase())};
    const b = [...document.querySelectorAll('.moduleMenu button')].find(x => x.textContent.replace(/\\d+$/,'').trim().toLowerCase() === want);
    if (!b) return false;
    const m = b.closest('.module'); if (m && !m.classList.contains('open') && typeof toggleModule === 'function') toggleModule(m.id);
    ['bc','sc','da','pm','al','ew'].forEach(k => { if (document.getElementById(k) && !b.dataset[k] && typeof window[k + 'Close_'] === 'function') window[k + 'Close_'](); });
    b.click(); return true;
  })()`;
  dash.webContents.executeJavaScript(js, true).then(() => setTimeout(probe, 300)).catch(() => {});
}

/* ---------- title bar buttons ---------- */
ipcMain.on('tb', (e, act, arg) => {
  if (!win) return;
  const wc = dash && dash.webContents;
  switch (act) {
    case 'min': win.minimize(); break;
    case 'max': win.isMaximized() ? win.unmaximize() : win.maximize(); break;
    case 'close': win.close(); break;
    case 'reload': wc && wc.reload(); break;
    case 'home': openPage('Overview'); break;
    case 'back': if (histPos > 0) { histPos--; histJump = true; openPage(history[histPos]); } break;
    case 'fwd': if (histPos < history.length - 1) { histPos++; histJump = true; openPage(history[histPos]); } break;
    case 'open': openPage(arg); break;
    case 'update': if (updater) { quitting = true; updater.quitAndInstall(false, true); } break;
    case 'badge': setBadgeImage(arg); break;
    case 'ready': push(); break;
  }
});

/* ---------- 6 · taskbar badge ---------- */
function setBadge(n) {
  if (!win) return;
  if (!n) { win.setOverlayIcon(null, ''); return; }
  win.webContents.send('make-badge', n);     // the title bar draws the red number, sends it back
}
function setBadgeImage(dataUrl) {
  if (win && dataUrl) win.setOverlayIcon(nativeImage.createFromDataURL(dataUrl), `${state.badge} discount approvals waiting`);
}

/* ---------- 5 · Windows notifications (from the activity log) ---------- */
function notify(title, body, go) {
  if (!Notification.isSupported()) return;
  const n = new Notification({ title, body, icon: ICON, silent: false });
  n.on('click', () => showMain(go));
  n.show();
}
const AUDIT_JS = `new Promise(res => {
  if (typeof AUTH_TOKEN === 'undefined' || !AUTH_TOKEN || !window.google || !google.script) return res(null);
  const me = (typeof AUTH_USER !== 'undefined' && AUTH_USER) ? String(AUTH_USER.name || AUTH_USER.username || '') : '';
  google.script.run.withSuccessHandler(d => res({ me, rows: ((d && d.rows) || []).slice(0, 60) }))
    .withFailureHandler(() => res(null)).getAuditLog(AUTH_TOKEN, { limit: 60 });
})`;
const ALERTS = [
  [/^Deal added/i, '🆕 New deal', 'Deals Details'],
  [/^Deal status changed/i, '🔄 Deal status changed', 'Deals Details'],
  [/^Payment recorded/i, '💵 Payment recorded', 'Collections'],
  [/^Payment deleted/i, '🗑 Payment deleted', 'Collections'],
  [/^Discount (approved|rejected)/i, '🏷 Discount decision', 'Discount Approvals'],
  [/^Broker commission paid/i, '🤝 Broker commission paid', 'Broker Commissions'],
  [/^Prices changed/i, '💲 Prices changed', 'Price Manager'],
  [/^Inventory edit/i, '✏️ Inventory edited in the sheet', 'Activity Log'],
  [/^Daily backup/i, '⚠️ Daily backup', 'Activity Log']
];
let seen = null;
async function pollAlerts() {
  if (!dash || dash.webContents.isLoading() || !state.loggedIn || state.offline) return;
  let r = null;
  try { r = await dash.webContents.executeJavaScript(AUDIT_JS, true); } catch (e) { return; }
  if (!r || !r.rows) return;
  const sig = x => [x.time, x.user, x.action, x.unit, x.client, x.details].join('|');
  if (!seen) seen = new Set(readJson('seen.json', []));
  const first = !seen.size, fresh = r.rows.filter(x => !seen.has(sig(x)));
  r.rows.forEach(x => seen.add(sig(x)));
  writeJson('seen.json', [...seen].slice(-400));
  if (first || !pref('alerts', true)) return;
  const mine = String(r.me || '').trim().toLowerCase();
  fresh.reverse().filter(x => String(x.user || '').trim().toLowerCase() !== mine).slice(0, 5).forEach(x => {
    const a = ALERTS.find(([re]) => re.test(x.action));
    if (!a) return;
    const what = [x.project, x.unit, x.client].filter(Boolean).join(' · ');
    notify(a[1], [what, x.details && x.details.slice(0, 140), x.user && 'by ' + x.user].filter(Boolean).join('\n'), a[2]);
  });
}
let lastBadge = null, loggedSince = 0;
function badgeAlert() {
  if (!state.loggedIn) { lastBadge = null; loggedSince = 0; return; }
  if (!loggedSince) loggedSince = Date.now();
  if (Date.now() - loggedSince < 30000) { lastBadge = state.badge; return; }   // data still loading
  if (lastBadge !== null && state.badge > lastBadge && pref('alerts', true))
    notify('🏷 Discount waiting for your approval', `${state.badge} deal${state.badge > 1 ? 's' : ''} above the discount limit`, 'Discount Approvals');
  lastBadge = state.badge;
}

/* ---------- 4 · tray icon + 7 · start with Windows ---------- */
function startWithWindows() { return app.getLoginItemSettings({ args: ['--hidden'] }).openAtLogin; }
function buildTray() {
  tray = new Tray(nativeImage.createFromPath(ICON).resize({ width: 16, height: 16 }));
  tray.setToolTip('EDEN BI');
  const menu = () => Menu.buildFromTemplate([
    { label: 'Open EDEN BI', click: () => showMain() },
    { type: 'separator' },
    ...['Overview', 'Add New Deal', 'Deals Details', 'Collections', 'Discount Approvals', 'Sales Scorecard', 'Broker Commissions']
      .map(p => ({ label: p, click: () => showMain(p) })),
    { type: 'separator' },
    { label: 'Notifications', type: 'checkbox', checked: pref('alerts', true), click: m => setPref('alerts', m.checked) },
    { label: 'Start with Windows', type: 'checkbox', checked: startWithWindows(),
      click: m => app.setLoginItemSettings({ openAtLogin: m.checked, args: ['--hidden'] }) },
    { label: 'Check for updates', click: () => checkUpdates(true) },
    { type: 'separator' },
    { label: 'Quit EDEN BI', click: () => { quitting = true; app.quit(); } }
  ]);
  tray.on('right-click', () => tray.popUpContextMenu(menu()));
  tray.on('click', () => showMain());
  tray.setContextMenu(null);
  if (process.platform !== 'win32') tray.setContextMenu(menu());
}

/* ---------- 8 · keyboard shortcuts ---------- */
const NUM_PAGES = ['Overview', 'Projects', 'Sales', 'Deals Details', 'Collections', 'Sales Scorecard', 'Discount Approvals', 'Broker Commissions', 'Activity Log'];
function shortcut(e, i) {
  if (i.type !== 'keyDown') return;
  const wc = dash && dash.webContents, ctrl = i.control || i.meta, k = (i.key || '').toLowerCase();
  const run = f => { e.preventDefault(); f(); };
  if (k === 'f5' || (ctrl && k === 'r')) return run(() => (i.shift ? wc.reloadIgnoringCache() : wc.reload()));
  if (k === 'f11') return run(() => win.setFullScreen(!win.isFullScreen()));
  if (i.alt && k === 'arrowleft') return run(() => ipcMain.emit('tb', {}, 'back'));
  if (i.alt && k === 'arrowright') return run(() => ipcMain.emit('tb', {}, 'fwd'));
  if (i.alt && k === 'home') return run(() => openPage('Overview'));
  if (!ctrl) return;
  if (k === 'n') return run(() => openPage('Add New Deal'));
  if (k === 'f') return run(() => wc.executeJavaScript(`(()=>{const s=document.getElementById('dsGlobalSearch');if(s){s.focus();s.select();}})()`).catch(() => {}));
  if (k === '=' || k === '+') return run(() => wc.setZoomLevel(Math.min(3, wc.getZoomLevel() + 0.5)));
  if (k === '-') return run(() => wc.setZoomLevel(Math.max(-3, wc.getZoomLevel() - 0.5)));
  if (k === '0') return run(() => wc.setZoomLevel(0));
  if (i.shift && k === 'i') return run(() => wc.openDevTools({ mode: 'detach' }));
  if (/^[1-9]$/.test(k)) return run(() => openPage(NUM_PAGES[+k - 1]));
}

/* ---------- 9 · self-update (from the "desktop-app" GitHub release) ---------- */
let updater = null;
function setupUpdates() {
  try { updater = require('electron-updater').autoUpdater; } catch (e) { return; }
  updater.autoDownload = true;
  updater.autoInstallOnAppQuit = true;
  updater.on('update-downloaded', info => {
    push({ update: info && info.version || 'new' });
    notify('EDEN BI update ready', 'Click "Restart to update" in the top bar (or it installs next time you open EDEN BI).');
  });
  updater.on('error', () => {});
  checkUpdates(false);
  setInterval(() => checkUpdates(false), 4 * 3600 * 1000);
}
function checkUpdates(manual) {
  if (!updater || !app.isPackaged) { if (manual) dialog.showMessageBox({ message: 'Updates work in the installed app only.' }); return; }
  updater.checkForUpdates().then(r => {
    if (manual && (!r || !r.updateInfo || r.updateInfo.version === app.getVersion()))
      dialog.showMessageBox({ title: 'EDEN BI', message: 'EDEN BI is up to date.', detail: 'Version ' + app.getVersion() });
  }).catch(() => { if (manual) dialog.showMessageBox({ message: 'Could not check for updates (no internet?).' }); });
}

/* ---------- start ---------- */
Menu.setApplicationMenu(null);
app.whenReady().then(() => {
  const ses = session.fromPartition(PART);
  handleDownloads(ses);
  offlineCache(ses);
  state.lastOnline = (readJson('online.json', {}) || {}).at || 0;
  if (!HIDDEN) createSplash();
  createMain();
  buildTray();
  setupUpdates();
  setInterval(probe, 1200);
  setInterval(() => { badgeAlert(); }, 5000);
  setInterval(pollAlerts, 60 * 1000);
  setTimeout(pollAlerts, 20 * 1000);
  setTimeout(() => { if (!shown) readyToShow(); }, 25000);   // never leave the splash up forever
});
app.on('before-quit', () => { quitting = true; });
global.__eden = { win: () => win, dash: () => dash, splash: () => splash, tray: () => tray, state: () => state, openPage, probe, pollAlerts, notify };
app.on('window-all-closed', () => { if (quitting) app.quit(); });
