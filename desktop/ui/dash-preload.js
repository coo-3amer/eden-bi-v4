// Runs inside the dashboard page (desktop app only).
// 1) window.edenDesktop: the page's server-call layer (shim.js) keeps a copy of
//    each "get…" answer here, and uses it when there is no internet.
// 2) Keeps the sign-in across app restarts (the page keeps it per tab only).
const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('edenDesktop', {
  app: 'windows',
  cacheGet: key => ipcRenderer.invoke('cache:get', String(key)),
  cacheSet: (key, text) => ipcRenderer.send('cache:set', String(key), String(text)),
  offline: off => ipcRenderer.send('net:state', !!off)
});
try {
  const K = ['EDEN_AUTH_TOKEN', 'EDEN_AUTH_USER'], B = k => 'EDEN_DESKTOP_' + k;
  if (!sessionStorage.getItem(K[0]) && localStorage.getItem(B(K[0])))
    K.forEach(k => { const v = localStorage.getItem(B(k)); if (v != null) sessionStorage.setItem(k, v); });
  setInterval(() => {
    K.forEach(k => {
      const v = sessionStorage.getItem(k);
      if (v == null) localStorage.removeItem(B(k)); else if (localStorage.getItem(B(k)) !== v) localStorage.setItem(B(k), v);
    });
  }, 1500);
} catch (e) {}
