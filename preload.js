const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('desktop', {
  platform: process.platform,
  setTheme: (t) => ipcRenderer.send('theme', t),
  reveal: (p) => ipcRenderer.send('reveal', p),
  officeConvert: (req) => ipcRenderer.invoke('office:convert', req),
  htmlToPdf: (html) => ipcRenderer.invoke('html:pdf', html),
  onSaved: (cb) => ipcRenderer.on('saved', (_e, p) => cb(p)),
});
