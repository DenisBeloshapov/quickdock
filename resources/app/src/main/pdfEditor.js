// src/main/pdfEditor.js — контекстное окно редактора PDF (reorder + rotate + delete страниц)
const { BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { PDFDocument, degrees } = require('pdf-lib');

let editorWindow = null;
let onSaveCallback = null;

function openPdfEditor(filePath, onSave) {
  if (!filePath || !fs.existsSync(filePath)) return false;
  onSaveCallback = onSave || null;

  if (editorWindow && !editorWindow.isDestroyed()) {
    editorWindow.show();
    editorWindow.focus();
    editorWindow.webContents.send('pdf:loaded', { path: filePath });
    return true;
  }

  editorWindow = new BrowserWindow({
    width: 1100, height: 720, minWidth: 900, minHeight: 600,
    title: 'Редактор PDF — QuickDock',
    backgroundColor: '#0A0A0A',
    autoHideMenuBar: true, show: false,
    webPreferences: {
      preload: path.join(__dirname, '..', 'preload', 'pdfEditor.js'),
      contextIsolation: true, nodeIntegration: false,
    },
  });
  editorWindow.setMenuBarVisibility(false);
  editorWindow.loadFile(path.join(__dirname, '..', 'renderer', 'pdfEditor.html'));
  editorWindow.once('ready-to-show', () => {
    editorWindow.show();
    editorWindow.webContents.send('pdf:loaded', { path: filePath });
  });
  editorWindow.on('closed', () => { editorWindow = null; onSaveCallback = null; });
  return true;
}

ipcMain.handle('pdf-editor:get-data', async (_e, filePath) => {
  try {
    if (!filePath || !fs.existsSync(filePath)) return { ok: false, error: 'Файл не найден' };
    const bytes = fs.readFileSync(filePath);
    const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
    const pages = doc.getPages();
    return {
      ok: true, path: filePath, pageCount: pages.length,
      pages: pages.map((p, i) => {
        const r = p.getRotation().angle || 0;
        const { width, height } = p.getSize();
        return { index: i, rotation: r, width, height };
      }),
    };
  } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
});

ipcMain.handle('pdf-editor:get-bytes', async (_e, filePath) => {
  try {
    if (!filePath || !fs.existsSync(filePath)) return { ok: false, error: 'Файл не найден' };
    const bytes = fs.readFileSync(filePath);
    return { ok: true, bytes: Array.from(bytes) };
  } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
});

ipcMain.handle('pdf-editor:save', async (_e, payload) => {
  try {
    const { path: filePath, order, rotations, saveAs } = payload;
    if (!filePath || !fs.existsSync(filePath)) return { ok: false, error: 'Файл не найден' };
    const srcDoc = await PDFDocument.load(fs.readFileSync(filePath), { ignoreEncryption: true });
    const outDoc = await PDFDocument.create();
    const indices = (Array.isArray(order) && order.length ? order : srcDoc.getPageIndices());
    const copied = await outDoc.copyPages(srcDoc, indices);
    copied.forEach((page, newPos) => {
      const oldIndex = indices[newPos];
      if (rotations && rotations[oldIndex] !== undefined && rotations[oldIndex] % 360 !== 0) {
        const cur = page.getRotation().angle || 0;
        page.setRotation(degrees((cur + rotations[oldIndex]) % 360));
      }
      outDoc.addPage(page);
    });
    const outBytes = await outDoc.save();
    let outPath;
    if (saveAs) {
      const r = await dialog.showSaveDialog(editorWindow, {
        title: 'Сохранить PDF как…',
        defaultPath: filePath,
        filters: [{ name: 'PDF', extensions: ['pdf'] }],
      });
      if (r.canceled || !r.filePath) return { ok: false, canceled: true };
      outPath = r.filePath;
      fs.writeFileSync(outPath, outBytes);
    } else {
      const dir = path.dirname(filePath);
      const ext = path.extname(filePath);
      const base = path.basename(filePath, ext);
      outPath = path.join(dir, `${base}-edited${ext}`);
      let n = 1;
      while (fs.existsSync(outPath)) { outPath = path.join(dir, `${base}-edited-${n}${ext}`); n++; }
      fs.writeFileSync(outPath, outBytes);
    }
    if (typeof onSaveCallback === 'function') onSaveCallback(outPath);
    return { ok: true, path: outPath };
  } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
});

ipcMain.handle('pdf-editor:overwrite', async (_e, payload) => {
  try {
    const { path: filePath, order, rotations } = payload;
    if (!filePath || !fs.existsSync(filePath)) return { ok: false, error: 'Файл не найден' };
    const srcDoc = await PDFDocument.load(fs.readFileSync(filePath), { ignoreEncryption: true });
    const outDoc = await PDFDocument.create();
    const indices = (Array.isArray(order) && order.length ? order : srcDoc.getPageIndices());
    const copied = await outDoc.copyPages(srcDoc, indices);
    copied.forEach((page, newPos) => {
      const oldIndex = indices[newPos];
      if (rotations && rotations[oldIndex] !== undefined && rotations[oldIndex] % 360 !== 0) {
        const cur = page.getRotation().angle || 0;
        page.setRotation(degrees((cur + rotations[oldIndex]) % 360));
      }
      outDoc.addPage(page);
    });
    fs.writeFileSync(filePath, await outDoc.save());
    if (typeof onSaveCallback === 'function') onSaveCallback(filePath);
    return { ok: true, path: filePath };
  } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
});

ipcMain.on('pdf-editor:confirm', (e, message) => {
  if (!editorWindow || editorWindow.isDestroyed()) { e.returnValue = false; return; }
  const choice = dialog.showMessageBoxSync(editorWindow, {
    type: 'question', buttons: ['Отмена', 'OK'], defaultId: 1, cancelId: 0,
    title: 'Подтверждение', message: message || 'Подтвердите действие',
  });
  e.returnValue = (choice === 1);
});

module.exports = { openPdfEditor };
