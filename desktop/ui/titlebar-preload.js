const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('tb', {
  act: (a, x) => ipcRenderer.send('tb', a, x),
  onState: fn => ipcRenderer.on('state', (e, s) => fn(s)),
  onBadge: fn => ipcRenderer.on('make-badge', (e, n) => fn(n))
});
