// src/preload/pdfEditor.js — preload для контекстного окна редактора PDF
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('qdPdfEditor', {
  getData: (filePath) => ipcRenderer.invoke('pdf-editor:get-data', filePath),
  getBytes: (filePath) => ipcRenderer.invoke('pdf-editor:get-bytes', filePath),
  save: (payload) => ipcRenderer.invoke('pdf-editor:save', payload),
  overwrite: (payload) => ipcRenderer.invoke('pdf-editor:overwrite', payload),
  onLoaded: (cb) => {
    const l = (_e, p) => cb(p);
    ipcRenderer.on('pdf:loaded', l);
    return () => ipcRenderer.removeListener('pdf:loaded', l);
  },
  confirm: (message) => ipcRenderer.sendSync('pdf-editor:confirm', message),
});
