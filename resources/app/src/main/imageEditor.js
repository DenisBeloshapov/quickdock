// src/main/imageEditor.js — контекстное окно редактора изображений (обрезка + поворот + порядок → PDF)
const { BrowserWindow, ipcMain, dialog, app } = require('electron');
const path = require('path');
const fs = require('fs');
const { PDFDocument } = require('pdf-lib');

const A4_W = 595.28;
const A4_H = 841.89;
const A4_MARGIN = 24;

let editorWindow = null;
let onSaveCallback = null;
let staged = []; // страницы текущего экспорта: { bytes, mime }

const IMAGE_EXTS = ['jpg', 'jpeg', 'jfif', 'png', 'bmp', 'gif', 'webp'];

// Собирает PDF из готовых (уже обрезанных/повёрнутых) картинок.
// pages: [{ bytes: Uint8Array|Buffer, mime: 'image/jpeg'|'image/png' }]
// mode: 'a4' — A4 с полями (ориентация по снимку), 'fit' — страница по пропорциям снимка, без полей.
async function buildPdf(pages, mode) {
  const pdf = await PDFDocument.create();
  for (let i = 0; i < pages.length; i++) {
    const pg = pages[i];
    const bytes = Buffer.from(pg.bytes);
    let img;
    try {
      img = pg.mime === 'image/jpeg' ? await pdf.embedJpg(bytes) : await pdf.embedPng(bytes);
    } catch (err) {
      // Нестандартный JPEG/PNG, который не принял pdf-lib: перекодируем средствами Electron.
      const { nativeImage } = require('electron');
      const ni = nativeImage.createFromBuffer(bytes);
      if (ni.isEmpty()) throw new Error('Не удалось прочитать изображение №' + (i + 1));
      img = await pdf.embedPng(ni.toPNG());
    }
    const w = img.width, h = img.height;
    if (mode === 'fit') {
      const k = A4_H / Math.max(w, h);
      const pw = w * k, ph = h * k;
      pdf.addPage([pw, ph]).drawImage(img, { x: 0, y: 0, width: pw, height: ph });
    } else {
      const landscape = w > h;
      const pw = landscape ? A4_H : A4_W;
      const ph = landscape ? A4_W : A4_H;
      const k = Math.min((pw - A4_MARGIN * 2) / w, (ph - A4_MARGIN * 2) / h);
      const dw = w * k, dh = h * k;
      pdf.addPage([pw, ph]).drawImage(img, { x: (pw - dw) / 2, y: (ph - dh) / 2, width: dw, height: dh });
    }
  }
  return pdf.save();
}

function openImageEditor(filePaths, onSave) {
  const list = (Array.isArray(filePaths) ? filePaths : [filePaths]).filter((p) => p && fs.existsSync(p));
  if (!list.length) return false;
  onSaveCallback = onSave || null;

  if (editorWindow && !editorWindow.isDestroyed()) {
    editorWindow.show();
    editorWindow.focus();
    editorWindow.webContents.send('img:add', { paths: list });
    return true;
  }

  editorWindow = new BrowserWindow({
    width: 1180, height: 760, minWidth: 820, minHeight: 560,
    title: 'Редактор изображений — QuickDock',
    backgroundColor: '#0A0A0A',
    autoHideMenuBar: true, show: false,
    webPreferences: {
      preload: path.join(__dirname, '..', 'preload', 'imageEditor.js'),
      contextIsolation: true, nodeIntegration: false,
    },
  });
  editorWindow.setMenuBarVisibility(false);
  editorWindow.loadFile(path.join(__dirname, '..', 'renderer', 'imageEditor.html'));
  // Отправляем после полной загрузки страницы, когда слушатель в renderer уже зарегистрирован.
  editorWindow.webContents.once('did-finish-load', () => {
    editorWindow.show();
    editorWindow.webContents.send('img:add', { paths: list });
  });
  editorWindow.on('closed', () => { editorWindow = null; onSaveCallback = null; staged = []; });
  return true;
}

ipcMain.handle('img-editor:read', async (_e, filePath) => {
  try {
    if (!filePath || !fs.existsSync(filePath)) return { ok: false, error: 'Файл не найден' };
    return { ok: true, name: path.basename(filePath), bytes: fs.readFileSync(filePath) };
  } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
});

// Кнопка «Обзор»: выбрать несколько файлов-изображений.
ipcMain.handle('img-editor:pick', async () => {
  try {
    const r = await dialog.showOpenDialog(editorWindow, {
      title: 'Добавить фото',
      properties: ['openFile', 'multiSelections'],
      filters: [{ name: 'Изображения', extensions: IMAGE_EXTS }],
    });
    if (r.canceled) return { ok: true, paths: [] };
    return { ok: true, paths: r.filePaths };
  } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
});

// Страницы приходят по одной, чтобы не гонять через IPC один огромный объект.
ipcMain.handle('img-editor:stage-reset', () => { staged = []; return true; });
ipcMain.handle('img-editor:stage-page', (_e, page) => {
  if (!page || !page.bytes) return false;
  staged.push({ bytes: page.bytes, mime: page.mime });
  return true;
});

ipcMain.handle('img-editor:finish', async (_e, opts) => {
  try {
    if (!staged.length) return { ok: false, error: 'Нет изображений' };
    const mode = opts && opts.pageMode === 'fit' ? 'fit' : 'a4';
    const outBytes = await buildPdf(staged, mode);
    staged = [];
    const base = String((opts && opts.suggestedName) || 'images').replace(/[\\/:*?"<>|]+/g, '_');
    const r = await dialog.showSaveDialog(editorWindow, {
      title: 'Сохранить PDF как…',
      defaultPath: path.join(app.getPath('documents'), base + '.pdf'),
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
    });
    if (r.canceled || !r.filePath) return { ok: false, canceled: true };
    fs.writeFileSync(r.filePath, outBytes);
    if (typeof onSaveCallback === 'function') onSaveCallback(r.filePath);
    return { ok: true, path: r.filePath };
  } catch (err) { staged = []; return { ok: false, error: String(err && err.message || err) }; }
});

ipcMain.on('img-editor:confirm', (e, message) => {
  if (!editorWindow || editorWindow.isDestroyed()) { e.returnValue = false; return; }
  const choice = dialog.showMessageBoxSync(editorWindow, {
    type: 'question', buttons: ['Отмена', 'OK'], defaultId: 1, cancelId: 0,
    title: 'Подтверждение', message: message || 'Подтвердите действие',
  });
  e.returnValue = (choice === 1);
});

module.exports = { openImageEditor, buildPdf };
