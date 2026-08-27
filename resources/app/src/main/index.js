const { app, BrowserWindow, Tray, Menu, screen, shell, nativeImage, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const store = require('./store');
const { ClipboardWatcher } = require('./clipboard');
const { buildPdf, classify } = require('./pdf');

app.commandLine.appendSwitch('high-dpi-support', 'true');
app.commandLine.appendSwitch('force-device-scale-factor', '1');

const SIZE_PRESETS = { small: 44, normal: 60, large: 76, xlarge: 92 };
const PANEL_W = 460, PANEL_H = 480;
const DOCK_MARGIN = 0;
const HOVER_LEAVE_DELAY = 350;
const PANEL_SHOW_DEBOUNCE = 60;
const ICON_RATIOS = { normal: 1.338, blink: 1.338, file: 1.455 };

let tray, iconWindow, panelWindow, clipboardWatcher, isQuitting = false;
let iconSize = 'normal', iconHidden = false, iconState = 'normal';
let iconHover = false, panelHover = false, panelVisible = false;
let hidePanelTimer = null, showPanelTimer = null, blinkTimer = null;
let filesList = [];
let isDragging = false;

function loadIconRatios() {
  try {
    for (const name of ['normal', 'blink', 'file']) {
      const ni = nativeImage.createFromPath(path.join(__dirname, '..', '..', 'build', `icon-${name}.png`));
      if (!ni.isEmpty()) {
        const sz = ni.getSize();
        if (sz.width > 0 && sz.height > 0) ICON_RATIOS[name] = sz.width / sz.height;
      }
    }
  } catch {}
}

function iconH() { return SIZE_PRESETS[iconSize]; }
function iconW(state = iconState) {
  const ratio = ICON_RATIOS[state] || ICON_RATIOS.normal;
  return Math.round(iconH() * ratio);
}
function iconBounds(state = iconState) { return { w: iconW(state), h: iconH() }; }

app.whenReady().then(async () => {
  if (!app.requestSingleInstanceLock()) { app.quit(); return; }
  loadIconRatios();
  const settings = store.loadSettings();
  iconSize = settings.iconSize || 'normal';
  iconHidden = !!settings.hidden;
  createTray();
  createIconWindow();
  createPanelWindow();
  setupIpc();
  clipboardWatcher = new ClipboardWatcher();
  clipboardWatcher.onNewText((history) => {
    if (panelWindow && !panelWindow.isDestroyed()) panelWindow.webContents.send('clipboard:updated', history);
    triggerBlink();
  });
  clipboardWatcher.start();
  positionIcon();
  if (!iconHidden) iconWindow.show();
  startCursorPolling();
});

app.on('second-instance', () => { if (iconWindow && !iconHidden) iconWindow.show(); });
app.on('window-all-closed', (e) => e.preventDefault());
app.on('activate', () => { if (iconWindow && !iconHidden) iconWindow.show(); });
app.on('before-quit', () => {
  isQuitting = true;
  try {
    if (clipboardWatcher) clipboardWatcher.stop();
    if (iconWindow && !iconWindow.isDestroyed()) { const b = iconWindow.getBounds(); store.saveSettings({ iconX: b.x, iconY: b.y }); }
    store.cleanTempDir();
  } catch {}
});

function createIconWindow() {
  const { w, h } = iconBounds();
  iconWindow = new BrowserWindow({
    width: w, height: h, frame: false, transparent: true,
    backgroundColor: '#00000000', resizable: false, maximizable: false,
    minimizable: false, fullscreenable: false, skipTaskbar: true,
    alwaysOnTop: true, hasShadow: false, show: false,
    webPreferences: { preload: path.join(__dirname, '..', 'preload', 'index.js'), contextIsolation: true, nodeIntegration: false },
  });
  iconWindow.setAlwaysOnTop(true, 'screen-saver');
  iconWindow.loadFile(path.join(__dirname, '..', 'renderer', 'icon.html'));
}

function createPanelWindow() {
  panelWindow = new BrowserWindow({
    width: PANEL_W, height: PANEL_H, frame: false, transparent: true,
    backgroundColor: '#00000000', resizable: false, maximizable: false,
    minimizable: false, fullscreenable: false, skipTaskbar: true,
    alwaysOnTop: true, hasShadow: false, show: true,
    webPreferences: { preload: path.join(__dirname, '..', 'preload', 'index.js'), contextIsolation: true, nodeIntegration: false },
  });
  panelWindow.setAlwaysOnTop(true, 'floating');
  panelWindow.loadFile(path.join(__dirname, '..', 'renderer', 'panel.html'));
  panelWindow.setIgnoreMouseEvents(true, { forward: true });
  panelWindow.once('did-finish-load', () => { repositionPanel(); panelWindow.webContents.send('panel:set-expanded', false); });
}

function positionIcon() {
  const settings = store.loadSettings();
  const { w, h } = iconBounds();
  const primary = screen.getPrimaryDisplay().workArea;
  if (typeof settings.iconX === 'number' && typeof settings.iconY === 'number') {
    let x = settings.iconX, y = settings.iconY;
    if (x + w > primary.x + primary.width) x = primary.x + primary.width - w - DOCK_MARGIN;
    if (x < primary.x) x = primary.x + DOCK_MARGIN;
    if (y + h > primary.y + primary.height) y = primary.y + primary.height - h - DOCK_MARGIN;
    if (y < primary.y) y = primary.y + DOCK_MARGIN;
    iconWindow.setPosition(x, y);
  } else {
    iconWindow.setPosition(primary.x + primary.width - w - DOCK_MARGIN, primary.y + primary.height - h - DOCK_MARGIN);
  }
}

function applyIconSize(newSize) {
  iconSize = newSize;
  const { w, h } = iconBounds();
  const bounds = iconWindow.getBounds();
  const bottomY = bounds.y + bounds.height;
  const centerX = bounds.x + Math.round(bounds.width / 2);
  iconWindow.setBounds({ x: centerX - Math.round(w / 2), y: bottomY - h, width: w, height: h });
  iconWindow.webContents.send('icon:set-size', iconH());
  if (panelVisible) repositionPanel();
}

function createTray() {
  let trayImg = nativeImage.createFromPath(path.join(__dirname, '..', '..', 'build', 'tray.png'));
  if (trayImg.isEmpty()) trayImg = nativeImage.createEmpty();
  tray = new Tray(trayImg);
  tray.setToolTip('QuickDock');
  rebuildTrayMenu();
}

function rebuildTrayMenu() {
  const labels = { small: 'Маленькая', normal: 'Обычная', large: 'Крупная', xlarge: 'Очень крупная' };
  const menu = Menu.buildFromTemplate([
    { label: iconHidden ? 'Показать' : 'Скрыть', click: () => toggleVisible() },
    { label: 'Настройки', submenu: [{ label: 'Размер иконки', submenu: Object.keys(SIZE_PRESETS).map((key) => ({
      label: labels[key], type: 'radio', checked: iconSize === key,
      click: () => { applyIconSize(key); store.saveSettings({ iconSize: key }); rebuildTrayMenu(); },
    })) }] },
    { type: 'separator' },
    { label: 'Поверх всех окон', click: () => forceOnTop() },
    { type: 'separator' },
    { label: 'Выход', click: () => { isQuitting = true; app.quit(); } },
  ]);
  tray.setContextMenu(menu);
}

function forceOnTop() {
  try {
    if (iconWindow && !iconWindow.isDestroyed()) { iconWindow.setAlwaysOnTop(false); iconWindow.setAlwaysOnTop(true, 'screen-saver'); iconWindow.moveTop(); }
    if (panelWindow && !panelWindow.isDestroyed()) { panelWindow.setAlwaysOnTop(false); panelWindow.setAlwaysOnTop(true, 'floating'); panelWindow.moveTop(); }
  } catch {}
}

function toggleVisible() {
  iconHidden = !iconHidden;
  store.saveSettings({ hidden: iconHidden });
  if (iconHidden) { iconWindow.hide(); hidePanelImmediate(); }
  else iconWindow.show();
  rebuildTrayMenu();
}

function showPanel() {
  if (!panelWindow || panelWindow.isDestroyed()) return;
  cancelHidePanel();
  if (panelVisible) return;
  if (showPanelTimer) clearTimeout(showPanelTimer);
  showPanelTimer = setTimeout(() => {
    showPanelTimer = null;
    if (panelVisible || (!iconHover && !panelHover)) return;
    panelVisible = true;
    repositionPanel();
    panelWindow.setIgnoreMouseEvents(false);
    panelWindow.moveTop();
    panelWindow.focus();
    if (iconWindow && !iconWindow.isDestroyed()) iconWindow.moveTop();
    panelWindow.webContents.send('clipboard:updated', clipboardWatcher.getHistory());
    panelWindow.webContents.send('files:added', filesList.slice());
    panelWindow.webContents.send('panel:set-expanded', true);
  }, PANEL_SHOW_DEBOUNCE);
}

function hidePanelImmediate() {
  if (showPanelTimer) { clearTimeout(showPanelTimer); showPanelTimer = null; }
  if (!panelWindow || panelWindow.isDestroyed() || !panelVisible) return;
  panelVisible = false;
  panelWindow.setIgnoreMouseEvents(true, { forward: true });
  panelWindow.webContents.send('panel:set-expanded', false);
}

function repositionPanel() {
  if (!iconWindow || !panelWindow) return;
  const iconB = iconWindow.getBounds();
  const wa = screen.getDisplayMatching(iconB).workArea;
  const iconCx = iconB.x + iconB.width / 2;
  let x, y;
  if (iconCx <= wa.x + wa.width / 2) { x = iconB.x; y = iconB.y + iconB.height - PANEL_H; }
  else { x = iconB.x + iconB.width - PANEL_W; y = iconB.y + iconB.height - PANEL_H; }
  if (x + PANEL_W > wa.x + wa.width) x = wa.x + wa.width - PANEL_W;
  if (x < wa.x) x = wa.x;
  if (y + PANEL_H > wa.y + wa.height) y = wa.y + wa.height - PANEL_H;
  if (y < wa.y) y = wa.y;
  panelWindow.setBounds({ x, y, width: PANEL_W, height: PANEL_H });
}

function scheduleHidePanel() {
  if (showPanelTimer) { clearTimeout(showPanelTimer); showPanelTimer = null; }
  if (hidePanelTimer) clearTimeout(hidePanelTimer);
  hidePanelTimer = setTimeout(() => { hidePanelTimer = null; if (!iconHover && !panelHover) hidePanelImmediate(); }, HOVER_LEAVE_DELAY);
}
function cancelHidePanel() { if (hidePanelTimer) { clearTimeout(hidePanelTimer); hidePanelTimer = null; } }

let cursorPollTimer = null;
function startCursorPolling() {
  if (cursorPollTimer) return;
  cursorPollTimer = setInterval(() => {
    if (!iconWindow || iconWindow.isDestroyed()) return;
    try {
      const cursor = screen.getCursorScreenPoint();
      const iconB = iconWindow.getBounds();
      const overIcon = cursor.x >= iconB.x && cursor.x < iconB.x + iconB.width && cursor.y >= iconB.y && cursor.y < iconB.y + iconB.height;
      let overPanel = false;
      if (panelWindow && !panelWindow.isDestroyed() && panelVisible) {
        const pb = panelWindow.getBounds();
        overPanel = cursor.x >= pb.x && cursor.x < pb.x + pb.width && cursor.y >= pb.y && cursor.y < pb.y + pb.height;
      }
      if (iconHover !== overIcon) { iconHover = overIcon; if (iconHover) { cancelHidePanel(); showPanel(); } else scheduleHidePanel(); }
      if (panelHover !== overPanel) { panelHover = overPanel; if (panelHover) cancelHidePanel(); else scheduleHidePanel(); }
    } catch {}
  }, 150);
}

function setIconState(state) {
  if (iconState === state) return;
  iconState = state;
  if (iconWindow && !iconWindow.isDestroyed()) iconWindow.webContents.send('icon:set-state', state);
}

function triggerBlink() {
  if (blinkTimer) clearTimeout(blinkTimer);
  const prevState = iconState;
  iconState = 'blink';
  if (iconWindow && !iconWindow.isDestroyed()) iconWindow.webContents.send('icon:set-state', 'blink');
  blinkTimer = setTimeout(() => { iconState = null; setIconState(prevState === 'blink' ? 'normal' : prevState); }, 180);
}

function moveIconBy(dx, dy) {
  if (!iconWindow) return;
  isDragging = true;
  const b = iconWindow.getBounds();
  let nx = b.x + dx, ny = b.y + dy;
  const wa = screen.getDisplayMatching(b).workArea;
  if (nx < wa.x) nx = wa.x;
  if (nx + b.width > wa.x + wa.width) nx = wa.x + wa.width - b.width;
  if (ny < wa.y) ny = wa.y;
  if (ny + b.height > wa.y + wa.height) ny = wa.y + wa.height - b.height;
  iconWindow.setPosition(nx, ny, false);
  if (panelVisible && panelWindow) repositionPanel();
}

function snapIconToEdge() {
  if (!iconWindow) return;
  isDragging = false;
  const bounds = iconWindow.getBounds();
  const wa = screen.getDisplayMatching(bounds).workArea;
  const cx = bounds.x + bounds.width / 2, cy = bounds.y + bounds.height / 2;
  const dLeft = cx - wa.x, dRight = wa.x + wa.width - cx, dTop = cy - wa.y, dBottom = wa.y + wa.height - cy;
  const minD = Math.min(dLeft, dRight, dTop, dBottom);
  let tx = bounds.x, ty = bounds.y;
  if (minD === dLeft) tx = wa.x + DOCK_MARGIN;
  else if (minD === dRight) tx = wa.x + wa.width - bounds.width - DOCK_MARGIN;
  else if (minD === dTop) ty = wa.y + DOCK_MARGIN;
  else ty = wa.y + wa.height - bounds.height - DOCK_MARGIN;
  if (minD === dLeft || minD === dRight) { if (ty + bounds.height > wa.y + wa.height) ty = wa.y + wa.height - bounds.height - DOCK_MARGIN; if (ty < wa.y) ty = wa.y + DOCK_MARGIN; }
  if (minD === dTop || minD === dBottom) { if (tx + bounds.width > wa.x + wa.width) tx = wa.x + wa.width - bounds.width - DOCK_MARGIN; if (tx < wa.x) tx = wa.x + DOCK_MARGIN; }
  const sx = bounds.x, sy = bounds.y, t0 = Date.now(), dur = 180;
  function step() {
    if (!iconWindow || iconWindow.isDestroyed()) return;
    const t = Math.min(1, (Date.now() - t0) / dur);
    const e = 1 - Math.pow(1 - t, 3);
    iconWindow.setPosition(Math.round(sx + (tx - sx) * e), Math.round(sy + (ty - sy) * e), false);
    if (panelVisible && panelWindow) repositionPanel();
    if (t < 1) setTimeout(step, 16);
    else { const fb = iconWindow.getBounds(); store.saveSettings({ iconX: fb.x, iconY: fb.y }); if (panelVisible && panelWindow) repositionPanel(); }
  }
  step();
}

function addFiles(paths) {
  for (const p of paths) {
    if (!p || typeof p !== 'string') continue;
    if (!fs.existsSync(p)) continue;
    const stat = fs.statSync(p);
    if (stat.isDirectory()) continue;
    const imported = store.importToTemp(p);
    filesList.push({ path: imported, name: path.basename(imported), size: stat.size, ext: path.extname(p).toLowerCase(), kind: classify(imported) });
  }
  if (panelWindow && !panelWindow.isDestroyed()) panelWindow.webContents.send('files:added', filesList.slice());
  return filesList.slice();
}

function setupIpc() {
  ipcMain.handle('clipboard:get', () => clipboardWatcher.getHistory());
  ipcMain.handle('clipboard:copy-back', (_e, text) => { clipboardWatcher.copyBack(text); triggerBlink(); });
  ipcMain.handle('files:add', (_e, paths) => addFiles(paths));
  ipcMain.handle('files:list', () => filesList.slice());
  ipcMain.handle('files:remove', (_e, p) => { filesList = filesList.filter((f) => f.path !== p); try { if (fs.existsSync(p)) fs.unlinkSync(p); } catch {} return filesList.slice(); });
  ipcMain.handle('files:clear', () => { for (const f of filesList) { try { if (fs.existsSync(f.path)) fs.unlinkSync(f.path); } catch {} } filesList = []; return filesList.slice(); });
  ipcMain.handle('files:build-pdf', async (_e, paths) => {
    let outName;
    if (paths && paths.length > 0) { outName = path.basename(paths[0]).replace(/\.[^.]+$/, '') + '.pdf'; }
    else { outName = `QuickDock_${Date.now()}.pdf`; }
    let outPath = path.join(store.getTempDir(), outName);
    let i = 1;
    while (fs.existsSync(outPath)) { outPath = path.join(store.getTempDir(), outName.replace(/\.pdf$/, '') + ` (${i}).pdf`); i++; }
    const result = await buildPdf(paths, outPath);
    const sourceSet = new Set(paths);
    filesList = filesList.filter((f) => { if (sourceSet.has(f.path)) { try { if (fs.existsSync(f.path)) fs.unlinkSync(f.path); } catch {} return false; } return true; });
    filesList.push({ path: result.outPath, name: path.basename(result.outPath), size: fs.statSync(result.outPath).size, ext: '.pdf', kind: classify(result.outPath), isOutput: true, errors: result.errors });
    return { ok: true, outPath: result.outPath, errors: result.errors, files: filesList.slice() };
  });
  ipcMain.handle('files:open-in-explorer', (_e, p) => shell.showItemInFolder(p));
  ipcMain.on('files:drag-out-sync', (e, { filePath, iconPath }) => {
    if (!panelWindow || panelWindow.isDestroyed()) { e.returnValue = false; return; }
    try { panelWindow.webContents.startDrag({ file: filePath, icon: iconPath || path.join(__dirname, '..', '..', 'build', 'icon.png') }); e.returnValue = true; }
    catch (err) { e.returnValue = false; }
  });
  ipcMain.handle('replies:get', () => store.loadReplies());
  ipcMain.handle('replies:save', (_e, list) => { store.saveReplies(list); return true; });
  ipcMain.handle('settings:get', () => store.loadSettings());
  ipcMain.handle('settings:set', (_e, partial) => store.saveSettings(partial));
  ipcMain.on('icon:hover', (_e, val) => { iconHover = !!val; if (iconHover) { cancelHidePanel(); showPanel(); } else scheduleHidePanel(); });
  ipcMain.on('panel:hover', (_e, val) => { panelHover = !!val; if (panelHover) cancelHidePanel(); else scheduleHidePanel(); });
  ipcMain.on('icon:file-hover', (_e, val) => { if (val) setIconState('file'); else if (iconState === 'file') setIconState('normal'); });
  ipcMain.on('icon:drag-start', () => { isDragging = true; cancelHidePanel(); });
  ipcMain.on('icon:drag-move', (_e, { dx, dy }) => moveIconBy(dx, dy));
  ipcMain.on('icon:drag-end', () => snapIconToEdge());
  ipcMain.on('panel:focus', () => { if (panelWindow && !panelWindow.isDestroyed()) panelWindow.focus(); });
}
