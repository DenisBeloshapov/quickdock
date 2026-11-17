'use strict';
// Мини-парсер BIFF8 (.xls): достаёт то, чего не отдаёт библиотека xlsx (community) —
// шрифты, границы, выравнивание, заливки по ячейкам, высоты строк и картинки (логотипы, печати).
// Значения ячеек по-прежнему берутся из библиотеки xlsx, тут только оформление.
const fs = require('fs');
const CFB = require('cfb');

// Палитра Excel по умолчанию (icv 8..63).
const DEFAULT_PALETTE = [
  0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
  0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
  0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
  0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
  0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
  0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
  0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333
];

function hex(n) { return '#' + (n & 0xFFFFFF).toString(16).padStart(6, '0'); }

function readRecords(buf) {
  const recs = [];
  let pos = 0;
  while (pos + 4 <= buf.length) {
    const type = buf.readUInt16LE(pos);
    const len = buf.readUInt16LE(pos + 2);
    if (pos + 4 + len > buf.length) break;
    recs.push({ type, pos, data: buf.subarray(pos + 4, pos + 4 + len) });
    pos += 4 + len;
  }
  return recs;
}

// Escher: рекурсивный обход. onNode(rec{ver,inst,type,data}) для каждого узла.
function walkEscher(buf, onNode) {
  let pos = 0;
  while (pos + 8 <= buf.length) {
    const verInst = buf.readUInt16LE(pos);
    const ver = verInst & 0xF;
    const inst = verInst >> 4;
    const type = buf.readUInt16LE(pos + 2);
    const len = buf.readUInt32LE(pos + 4);
    const end = Math.min(pos + 8 + len, buf.length);
    const data = buf.subarray(pos + 8, end);
    const node = { ver, inst, type, data };
    onNode(node);
    if (ver === 0xF) walkEscher(data, onNode);
    pos = end;
  }
}

// Склеивает запись с идущими следом CONTINUE (0x3C).
function joinContinue(recs, i) {
  const parts = [recs[i].data];
  let j = i + 1;
  while (j < recs.length && recs[j].type === 0x3C) { parts.push(recs[j].data); j++; }
  return Buffer.concat(parts);
}

function extractBlips(groupBuf) {
  const blips = [];
  walkEscher(groupBuf, (n) => {
    if (n.type !== 0xF007) return; // BSE
    // 36 байт заголовка BSE + имя, дальше — встроенный BLIP.
    const cbName = n.data.length > 33 ? n.data[33] : 0;
    const blipStart = 36 + cbName;
    let img = null;
    if (n.data.length > blipStart + 8) {
      const head = n.data.subarray(blipStart, Math.min(blipStart + 8 + 64, n.data.length));
      const body = n.data.subarray(blipStart + 8);
      // Ищем сигнатуру JPEG / PNG в первых байтах (после UID).
      for (let k = 0; k < Math.min(48, body.length - 4); k++) {
        if (body[k] === 0xFF && body[k + 1] === 0xD8 && body[k + 2] === 0xFF) { img = { mime: 'image/jpeg', data: body.subarray(k) }; break; }
        if (body[k] === 0x89 && body[k + 1] === 0x50 && body[k + 2] === 0x4E && body[k + 3] === 0x47) { img = { mime: 'image/png', data: body.subarray(k) }; break; }
      }
      void head;
    }
    blips.push(img);
  });
  return blips;
}

function extractShapes(drawBuf) {
  // Возвращает [{ pib, anchor:{c1,dx1,r1,dy1,c2,dx2,r2,dy2} }]
  const out = [];
  let cur = null;
  walkEscher(drawBuf, (n) => {
    if (n.type === 0xF004) { cur = { pib: 0, anchor: null }; out.push(cur); return; }
    if (!cur) return;
    if (n.type === 0xF00B) { // Opt
      const count = n.inst;
      for (let k = 0; k < count && (k + 1) * 6 <= n.data.length; k++) {
        const id = n.data.readUInt16LE(k * 6) & 0x3FFF;
        const val = n.data.readUInt32LE(k * 6 + 2);
        if (id === 0x104) cur.pib = val; // pib — индекс картинки (1-based)
      }
    } else if (n.type === 0xF010 && n.data.length >= 18) { // ClientAnchor
      const d = n.data;
      cur.anchor = {
        c1: d.readUInt16LE(2), dx1: d.readUInt16LE(4), r1: d.readUInt16LE(6), dy1: d.readUInt16LE(8),
        c2: d.readUInt16LE(10), dx2: d.readUInt16LE(12), r2: d.readUInt16LE(14), dy2: d.readUInt16LE(16)
      };
    }
  });
  return out.filter(s => s.pib > 0 && s.anchor);
}

function parseXls(fp) {
  let cfb;
  try { cfb = CFB.read(fp, { type: 'file' }); } catch { return null; }
  const entry = CFB.find(cfb, 'Workbook') || CFB.find(cfb, 'Book');
  if (!entry || !entry.content) return null;
  const buf = Buffer.from(entry.content);
  const recs = readRecords(buf);

  const palette = DEFAULT_PALETTE.slice();
  const fonts = []; // индекс по номерам Excel (4 пропущен)
  const xfs = [];
  const boundSheets = []; // { pos, name }
  let blips = [];
  const sheets = []; // по порядку BoundSheet: { cellXf, rows, defaultH, shapes }

  const color = (icv, auto) => {
    if (icv >= 8 && icv <= 63) return hex(palette[icv - 8]);
    if (icv === 0) return '#000000'; if (icv === 1) return '#FFFFFF';
    if (icv === 2) return '#FF0000'; if (icv === 3) return '#00FF00';
    if (icv === 4) return '#0000FF'; if (icv === 5) return '#FFFF00';
    if (icv === 6) return '#FF00FF'; if (icv === 7) return '#00FFFF';
    return auto;
  };

  let section = 'globals';
  let sheetIdx = -1;
  let cur = null;
  let fontCount = 0;

  for (let i = 0; i < recs.length; i++) {
    const r = recs[i];
    const d = r.data;
    if (r.type === 0x809) { // BOF
      const dt = d.length >= 4 ? d.readUInt16LE(2) : 0;
      if (dt === 0x0010) { // worksheet
        section = 'sheet';
        let k = boundSheets.findIndex(b => b.pos === r.pos);
        if (k < 0) k = ++sheetIdx; else sheetIdx = k;
        cur = sheets[k] = { cellXf: new Map(), rows: new Map(), colXf: new Map(), defaultH: null, shapes: [] };
      } else if (dt !== 0x0005) { section = 'other'; cur = null; }
      continue;
    }
    if (r.type === 0x0A) { if (section !== 'globals') { section = 'globals_done'; cur = null; } continue; } // EOF

    if (section === 'globals') {
      if (r.type === 0x85 && d.length >= 8) { // BOUNDSHEET
        boundSheets.push({ pos: d.readUInt32LE(0), hidden: d[4] & 3, name: '' });
      } else if (r.type === 0x31 && d.length >= 16) { // FONT
        const cch = d[14];
        const flags = d[15];
        let name = '';
        try { name = (flags & 1) ? d.subarray(16, 16 + cch * 2).toString('utf16le') : d.subarray(16, 16 + cch).toString('latin1'); } catch {}
        const f = {
          size: d.readUInt16LE(0) / 20,
          italic: !!(d.readUInt16LE(2) & 2),
          strike: !!(d.readUInt16LE(2) & 8),
          color: d.readUInt16LE(4),
          bold: d.readUInt16LE(6) >= 700,
          underline: d[10] !== 0,
          name
        };
        fonts[fontCount < 4 ? fontCount : fontCount + 1] = f; // Excel пропускает индекс 4
        fontCount++;
      } else if (r.type === 0x92 && d.length >= 2) { // PALETTE
        const n = d.readUInt16LE(0);
        for (let k = 0; k < n && 2 + k * 4 + 3 <= d.length && k < 56; k++) {
          palette[k] = (d[2 + k * 4] << 16) | (d[3 + k * 4] << 8) | d[4 + k * 4];
        }
      } else if (r.type === 0xE0 && d.length >= 20) { // XF
        const b = d.readUInt16LE(10);
        const c1 = d.readUInt16LE(12);
        const c2 = d.readUInt32LE(14);
        const c3 = d.readUInt16LE(18);
        xfs.push({
          font: d.readUInt16LE(0),
          halign: d[6] & 7,
          wrap: !!(d[6] & 8),
          valign: (d[6] >> 4) & 7,
          indent: d[8] & 0xF,
          bL: b & 0xF, bR: (b >> 4) & 0xF, bT: (b >> 8) & 0xF, bB: (b >> 12) & 0xF,
          cL: c1 & 0x7F, cR: (c1 >> 7) & 0x7F,
          cT: c2 & 0x7F, cB: (c2 >> 7) & 0x7F,
          pattern: (c2 >>> 26) & 0x3F,
          fore: c3 & 0x7F, back: (c3 >> 7) & 0x7F
        });
      } else if (r.type === 0xEB) { // MSODRAWINGGROUP
        try { blips = extractBlips(joinContinue(recs, i)); } catch {}
      }
      continue;
    }

    if (section === 'sheet' && cur) {
      switch (r.type) {
        case 0x201: case 0x27E: case 0x203: case 0x204: case 0xFD: case 0x205: case 0x06:
          if (d.length >= 6) cur.cellXf.set(d.readUInt16LE(0) + ',' + d.readUInt16LE(2), d.readUInt16LE(4));
          break;
        case 0xBD: { // MULRK
          if (d.length < 6) break;
          const row = d.readUInt16LE(0), c0 = d.readUInt16LE(2);
          const n = (d.length - 6) / 6 | 0;
          for (let k = 0; k < n; k++) cur.cellXf.set(row + ',' + (c0 + k), d.readUInt16LE(4 + k * 6));
          break;
        }
        case 0xBE: { // MULBLANK
          if (d.length < 6) break;
          const row = d.readUInt16LE(0), c0 = d.readUInt16LE(2);
          const n = (d.length - 6) / 2 | 0;
          for (let k = 0; k < n; k++) cur.cellXf.set(row + ',' + (c0 + k), d.readUInt16LE(4 + k * 2));
          break;
        }
        case 0x208: { // ROW
          if (d.length < 16) break;
          const h = d.readUInt16LE(6);
          const grbit = d.readUInt16LE(12);
          cur.rows.set(d.readUInt16LE(0), { h: (h & 0x8000) ? null : (h & 0x7FFF) / 20, hidden: !!(grbit & 0x20), xf: (grbit & 0x80) ? (d.readUInt16LE(14) & 0xFFF) : null });
          break;
        }
        case 0x7D: { // COLINFO — стиль столбца по умолчанию (для ячеек без собственной записи)
          if (d.length < 8) break;
          const c0 = d.readUInt16LE(0), c1 = Math.min(d.readUInt16LE(2), 255);
          const xf = d.readUInt16LE(6);
          for (let k = c0; k <= c1; k++) cur.colXf.set(k, xf);
          break;
        }
        case 0x225: // DEFAULTROWHEIGHT
          if (d.length >= 4) cur.defaultH = (d.readUInt16LE(2) & 0x7FFF) / 20;
          break;
        case 0xEC: // MSODRAWING
          try { cur.shapes.push(...extractShapes(joinContinue(recs, i))); } catch {}
          break;
        default: break;
      }
    }
  }

  const BORDER = { 0: null, 1: '1px solid', 2: '2px solid', 3: '1px dashed', 4: '1px dotted', 5: '3px solid', 6: '3px double', 7: '1px dotted', 8: '2px dashed', 9: '1px dashed', 10: '2px dashed', 11: '1px dashed', 12: '2px dashed', 13: '2px dashed' };
  const styles = xfs.map(x => {
    const f = fonts[x.font] || fonts[0] || { size: 10, name: 'Arial', color: 64 };
    const bd = (dg, icv) => (BORDER[dg] ? `${BORDER[dg]} ${color(icv, '#000000')}` : null);
    return {
      font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, underline: f.underline, strike: f.strike, color: color(f.color, '#000000') },
      halign: x.halign, valign: x.valign, wrap: x.wrap, indent: x.indent,
      border: { l: bd(x.bL, x.cL), r: bd(x.bR, x.cR), t: bd(x.bT, x.cT), b: bd(x.bB, x.cB) },
      fill: x.pattern === 1 && x.fore !== 64 ? color(x.fore, null) : null
    };
  });

  return { styles, blips, sheets, defaultFont: fonts[0] || null };
}

module.exports = { parseXls };
