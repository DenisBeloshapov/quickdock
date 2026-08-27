const { contextBridge, ipcRenderer, webUtils } = require('electron');
contextBridge.exposeInMainWorld('qd', {
  clipboard: {
    get: () => ipcRenderer.invoke('clipboard:get'),
    copyBack: (t) => ipcRenderer.invoke('clipboard:copy-back', t),
    onUpdated: (cb) => { const l = (_e, h) => cb(h); ipcRenderer.on('clipboard:updated', l); return () => ipcRenderer.removeListener('clipboard:updated', l); },
  },
  files: {
    add: (p) => ipcRenderer.invoke('files:add', p),
    list: () => ipcRenderer.invoke('files:list'),
    remove: (p) => ipcRenderer.invoke('files:remove', p),
    clear: () => ipcRenderer.invoke('files:clear'),
    buildPdf: (p) => ipcRenderer.invoke('files:build-pdf', p),
    openInExplorer: (p) => ipcRenderer.invoke('files:open-in-explorer', p),
    dragOut: (filePath, iconPath) => ipcRenderer.sendSync('files:drag-out-sync', { filePath, iconPath }),
    onAdded: (cb) => { const l = (_e, f) => cb(f); ipcRenderer.on('files:added', l); return () => ipcRenderer.removeListener('files:added', l); },
    pathForFile: (file) => { try { return webUtils.getPathForFile(file); } catch { return ''; } },
  },
  replies: { get: () => ipcRenderer.invoke('replies:get'), save: (l) => ipcRenderer.invoke('replies:save', l) },
  settings: { get: () => ipcRenderer.invoke('settings:get'), set: (p) => ipcRenderer.invoke('settings:set', p) },
  icon: {
    setHover: (v) => ipcRenderer.send('icon:hover', v),
    setFileHover: (v) => ipcRenderer.send('icon:file-hover', v),
    dragStart: () => ipcRenderer.send('icon:drag-start'),
    dragMove: (dx, dy) => ipcRenderer.send('icon:drag-move', { dx, dy }),
    dragEnd: () => ipcRenderer.send('icon:drag-end'),
    onSetState: (cb) => { const l = (_e, s) => cb(s); ipcRenderer.on('icon:set-state', l); return () => ipcRenderer.removeListener('icon:set-state', l); },
    onSetSize: (cb) => { const l = (_e, px) => cb(px); ipcRenderer.on('icon:set-size', l); return () => ipcRenderer.removeListener('icon:set-size', l); },
  },
  panel: {
    setHover: (v) => ipcRenderer.send('panel:hover', v),
    focus: () => ipcRenderer.send('panel:focus'),
    onSetExpanded: (cb) => { const l = (_e, v) => cb(v); ipcRenderer.on('panel:set-expanded', l); return () => ipcRenderer.removeListener('panel:set-expanded', l); },
  },
});
