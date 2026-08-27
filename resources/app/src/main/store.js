const fs = require('fs');
const path = require('path');
const { app } = require('electron');
let _appDir = null;
function getAppDir() { if (_appDir) return _appDir; _appDir = app.isPackaged ? path.dirname(app.getPath('exe')) : path.resolve(__dirname, '..', '..'); return _appDir; }
function getTempDir() { const dir = path.join(getAppDir(), 'quickdock-temp'); if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true }); return dir; }
function getRepliesPath() { return path.join(getAppDir(), 'quickdock-replies.json'); }
function getSettingsPath() { return path.join(getAppDir(), 'quickdock-settings.json'); }
function loadReplies() { try { const p = JSON.parse(fs.readFileSync(getRepliesPath(), 'utf8')); return Array.isArray(p) ? p : []; } catch { return []; } }
function saveReplies(r) { fs.writeFileSync(getRepliesPath(), JSON.stringify(r, null, 2), 'utf8'); }
function loadSettings() { try { return { iconSize: 'normal', hidden: false, ...JSON.parse(fs.readFileSync(getSettingsPath(), 'utf8')) }; } catch { return { iconSize: 'normal', hidden: false }; } }
function saveSettings(s) { const m = { ...loadSettings(), ...s }; fs.writeFileSync(getSettingsPath(), JSON.stringify(m, null, 2), 'utf8'); return m; }
function cleanTempDir() { try { const d = getTempDir(); if (fs.existsSync(d)) for (const e of fs.readdirSync(d)) try { fs.unlinkSync(path.join(d, e)); } catch {} } catch {} }
function importToTemp(src) { const base = path.basename(src); let final = path.join(getTempDir(), base); let i = 1; while (fs.existsSync(final)) { const ext = path.extname(base); const stem = path.basename(base, ext); final = path.join(getTempDir(), `${stem} (${i})${ext}`); i++; } fs.copyFileSync(src, final); return final; }
module.exports = { getAppDir, getTempDir, getRepliesPath, getSettingsPath, loadReplies, saveReplies, loadSettings, saveSettings, cleanTempDir, importToTemp };
