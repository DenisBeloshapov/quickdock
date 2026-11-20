// src/preload/imageEditor.js — preload для контекстного окна редактора изображений
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('qdImageEditor', {
  read: (filePath) => ipcRenderer.invoke('img-editor:read', filePath),
  pick: () => ipcRenderer.invoke('img-editor:pick'),
  stageReset: () => ipcRenderer.invoke('img-editor:stage-reset'),
  stagePage: (page) => ipcRenderer.invoke('img-editor:stage-page', page),
  finish: (opts) => ipcRenderer.invoke('img-editor:finish', opts),
  onAdd: (cb) => {
    const l = (_e, p) => cb(p);
    ipcRenderer.on('img:add', l);
    return () => ipcRenderer.removeListener('img:add', l);
  },
  confirm: (message) => ipcRenderer.sendSync('img-editor:confirm', message),
});
