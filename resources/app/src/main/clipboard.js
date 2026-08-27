const { clipboard } = require('electron');
const POLL_MS = 500, HISTORY_LIMIT = 5;
class ClipboardWatcher {
  constructor() { this.timer = null; this.history = []; this.listeners = []; }
  start() { const init = this._read(); if (init) this.history = [init]; this.timer = setInterval(() => this._tick(), POLL_MS); }
  stop() { if (this.timer) clearInterval(this.timer); this.timer = null; }
  onNewText(cb) { this.listeners.push(cb); }
  _read() { try { const t = clipboard.readText(); if (t && t.trim().length > 0) return t; } catch {} return null; }
  _tick() { const t = this._read(); if (!t || this.history[0] === t) return; this.history = [t, ...this.history.filter((h) => h !== t)].slice(0, HISTORY_LIMIT); for (const cb of this.listeners) cb(this.history.slice()); }
  copyBack(t) { clipboard.writeText(t); this.history = [t, ...this.history.filter((h) => h !== t)].slice(0, HISTORY_LIMIT); }
  getHistory() { return this.history.slice(); }
}
module.exports = { ClipboardWatcher };
