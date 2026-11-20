'use strict';
// Разбор оформления .xlsx (шрифты, заливки, границы, высоты строк, стили строк/столбцов, картинки).
// Возвращает ту же структуру, что и parseXls, чтобы рендер таблицы не зависел от формата.
// Значения ячеек по-прежнему берутся из библиотеки xlsx.
const fs = require('fs');
const JSZip = require('jszip');
const { DOMParser } = require('@xmldom/xmldom');

const INDEXED = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'
];

const BORDER = {
  thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed',
  dashDot: '1px dashed', dashDotDot: '1px dashed', mediumDashed: '2px dashed',
  mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed',
  medium: '2px solid', thick: '3px solid', double: '3px double'
};

const HALIGN = { general: 0, left: 1, center: 2, right: 3, fill: 4, justify: 5, centerContinuous: 6, distributed: 7 };
const VALIGN = { top: 0, center: 1, bottom: 2, justify: 3, distributed: 4 };

const parseXml = (s) => new DOMParser({ errorHandler: () => {} }).parseFromString(s, 'text/xml');
const kids = (node, name) => {
  const out = [];
  if (!node) return out;
  for (let c = node.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && (!name || c.localName === name)) out.push(c);
  }
  return out;
};
const kid = (node, name) => kids(node, name)[0] || null;
const attr = (node, name, def = null) => (node && node.hasAttribute(name) ? node.getAttribute(name) : def);

function applyTint(hexColor, tint) {
  if (!tint) return hexColor;
  const n = parseInt(hexColor.slice(1), 16);
  let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  // HLS-подобная коррекция как в Excel (упрощённо — линейное осветление/затемнение).
  const f = (v) => Math.round(tint < 0 ? v * (1 + tint) : v * (1 - tint) + 255 * tint);
  r = f(r); g = f(g); b = f(b);
  return '#' + ((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1);
}

function resolveRel(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/'); parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop(); else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

async function parseXlsx(fp) {
  const zip = await JSZip.loadAsync(fs.readFileSync(fp));
  const text = async (name) => { const f = zip.file(name); return f ? f.async('string') : null; };

  // --- Тема (цвета) ---
  const themeColors = [];
  const themeXml = await text('xl/theme/theme1.xml');
  if (themeXml) {
    const doc = parseXml(themeXml);
    const scheme = doc.getElementsByTagName('a:clrScheme')[0];
    if (scheme) {
      const byName = {};
      for (const c of kids(scheme)) {
        const v = kid(c, 'srgbClr') ? attr(kid(c, 'srgbClr'), 'val') : kid(c, 'sysClr') ? attr(kid(c, 'sysClr'), 'lastClr', '000000') : '000000';
        byName[c.localName] = '#' + v;
      }
      // Порядок индексов темы в Excel: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink
      ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
        .forEach(n => themeColors.push(byName[n] || '#000000'));
    }
  }
  const color = (node, auto) => {
    if (!node) return auto;
    if (node.hasAttribute('rgb')) return '#' + node.getAttribute('rgb').slice(-6);
    if (node.hasAttribute('indexed')) { const i = +node.getAttribute('indexed'); return i < INDEXED.length ? INDEXED[i] : auto; }
    if (node.hasAttribute('theme')) {
      const c = themeColors[+node.getAttribute('theme')] || auto;
      return c ? applyTint(c, parseFloat(attr(node, 'tint', '0'))) : auto;
    }
    return auto;
  };

  // --- Стили ---
  const styles = [];
  const stylesXml = await text('xl/styles.xml');
  if (!stylesXml) return null;
  const sdoc = parseXml(stylesXml).documentElement;
  const fontsEl = kid(sdoc, 'fonts'), fillsEl = kid(sdoc, 'fills'), bordersEl = kid(sdoc, 'borders'), xfsEl = kid(sdoc, 'cellXfs');
  const fonts = kids(fontsEl, 'font').map(f => {
    const flag = (n) => { const e = kid(f, n); return !!e && attr(e, 'val', '1') !== '0' && attr(e, 'val', '1') !== 'false'; };
    const u = kid(f, 'u');
    return {
      name: attr(kid(f, 'name'), 'val', 'Calibri'),
      size: parseFloat(attr(kid(f, 'sz'), 'val', '11')),
      bold: flag('b'), italic: flag('i'), strike: flag('strike'),
      underline: !!u && attr(u, 'val', 'single') !== 'none',
      color: color(kid(f, 'color'), '#000000')
    };
  });
  const fills = kids(fillsEl, 'fill').map(f => {
    const p = kid(f, 'patternFill');
    if (p && attr(p, 'patternType') === 'solid') return color(kid(p, 'fgColor'), null);
    return null;
  });
  const borders = kids(bordersEl, 'border').map(b => {
    const side = (n) => {
      const e = kid(b, n);
      const st = e ? attr(e, 'style') : null;
      return st && BORDER[st] ? `${BORDER[st]} ${color(kid(e, 'color'), '#000000')}` : null;
    };
    return { l: side('left'), r: side('right'), t: side('top'), b: side('bottom') };
  });
  for (const xf of kids(xfsEl, 'xf')) {
    const al = kid(xf, 'alignment');
    styles.push({
      font: fonts[+attr(xf, 'fontId', '0')] || fonts[0] || { name: 'Calibri', size: 11, color: '#000000' },
      halign: al ? (HALIGN[attr(al, 'horizontal', 'general')] ?? 0) : 0,
      valign: al ? (VALIGN[attr(al, 'vertical', 'bottom')] ?? 2) : 2,
      wrap: !!al && attr(al, 'wrapText') === '1' || !!al && attr(al, 'wrapText') === 'true',
      indent: al ? +attr(al, 'indent', '0') : 0,
      border: borders[+attr(xf, 'borderId', '0')] || { l: null, r: null, t: null, b: null },
      fill: fills[+attr(xf, 'fillId', '0')] || null
    });
  }

  // --- Листы в порядке workbook.xml ---
  const wbXml = await text('xl/workbook.xml');
  const wbRels = await text('xl/_rels/workbook.xml.rels');
  if (!wbXml || !wbRels) return { styles, blips: [], sheets: [] };
  const relMap = {};
  for (const r of kids(parseXml(wbRels).documentElement, 'Relationship')) relMap[attr(r, 'Id')] = attr(r, 'Target');
  const sheetEls = kids(kid(parseXml(wbXml).documentElement, 'sheets'), 'sheet');

  const blips = []; // общий список картинок; shape.pib = индекс+1
  const blipIndex = new Map();
  const mimeOf = (n) => /\.png$/i.test(n) ? 'image/png' : /\.jpe?g$/i.test(n) ? 'image/jpeg' : /\.gif$/i.test(n) ? 'image/gif' : /\.bmp$/i.test(n) ? 'image/bmp' : /\.svg$/i.test(n) ? 'image/svg+xml' : null;
  const getBlip = async (path) => {
    if (blipIndex.has(path)) return blipIndex.get(path);
    const mime = mimeOf(path); const f = zip.file(path);
    let idx = 0;
    if (mime && f) { blips.push({ mime, data: await f.async('nodebuffer') }); idx = blips.length; }
    blipIndex.set(path, idx);
    return idx;
  };

  const EMU = 9525; // EMU на пиксель
  const sheets = [];

  for (const sel of sheetEls) {
    const rid = sel.getAttribute('r:id') || sel.getAttribute('id');
    const target = relMap[rid];
    const out = { cellXf: new Map(), rows: new Map(), colXf: new Map(), cols: new Map(), defaultH: null, shapes: [] };
    sheets.push(out);
    if (!target) continue;
    const sheetPath = resolveRel('xl/workbook.xml', target);
    const sx = await text(sheetPath);
    if (!sx) continue;
    const sdocEl = parseXml(sx).documentElement;

    const fmt = kid(sdocEl, 'sheetFormatPr');
    if (fmt && fmt.hasAttribute('defaultRowHeight')) out.defaultH = parseFloat(fmt.getAttribute('defaultRowHeight'));

    for (const c of kids(kid(sdocEl, 'cols'), 'col')) {
      const c0 = +c.getAttribute('min') - 1, c1 = Math.min(+c.getAttribute('max') - 1, 255);
      const st = c.hasAttribute('style') ? +c.getAttribute('style') : null;
      const w = c.hasAttribute('width') ? parseFloat(c.getAttribute('width')) : null;
      const hidden = attr(c, 'hidden') === '1' || attr(c, 'hidden') === 'true';
      for (let k = c0; k <= c1; k++) { if (st != null) out.colXf.set(k, st); out.cols.set(k, { w, hidden }); }
    }

    for (const row of kids(kid(sdocEl, 'sheetData'), 'row')) {
      const r = +row.getAttribute('r') - 1;
      const custom = attr(row, 'customHeight') === '1' || attr(row, 'customHeight') === 'true';
      out.rows.set(r, {
        h: row.hasAttribute('ht') ? parseFloat(row.getAttribute('ht')) : null,
        hidden: attr(row, 'hidden') === '1' || attr(row, 'hidden') === 'true',
        xf: (attr(row, 'customFormat') === '1' || attr(row, 'customFormat') === 'true') && row.hasAttribute('s') ? +row.getAttribute('s') : null,
        custom
      });
      for (const c of kids(row, 'c')) {
        const ref = c.getAttribute('r');
        const m = /^([A-Z]+)(\d+)$/.exec(ref || '');
        if (!m) continue;
        let col = 0; for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
        out.cellXf.set(`${+m[2] - 1},${col - 1}`, c.hasAttribute('s') ? +c.getAttribute('s') : 0);
      }
    }

    // --- Картинки листа ---
    const relsPath = sheetPath.replace(/([^/]+)$/, '_rels/$1.rels');
    const relsXml = await text(relsPath);
    if (!relsXml) continue;
    for (const rel of kids(parseXml(relsXml).documentElement, 'Relationship')) {
      if (!/\/drawing$/.test(attr(rel, 'Type', ''))) continue;
      const drawPath = resolveRel(sheetPath, attr(rel, 'Target'));
      const dx = await text(drawPath);
      if (!dx) continue;
      const drelsXml = await text(drawPath.replace(/([^/]+)$/, '_rels/$1.rels'));
      const drel = {};
      if (drelsXml) for (const r2 of kids(parseXml(drelsXml).documentElement, 'Relationship')) drel[attr(r2, 'Id')] = resolveRel(drawPath, attr(r2, 'Target'));
      const ddoc = parseXml(dx).documentElement;
      for (const anchor of kids(ddoc)) {
        const kind = anchor.localName;
        if (!/Anchor$/.test(kind)) continue;
        const pic = kid(anchor, 'pic');
        if (!pic) continue;
        const blip = pic.getElementsByTagName('a:blip')[0];
        const embed = blip && (blip.getAttribute('r:embed') || blip.getAttribute('embed'));
        const path = embed && drel[embed];
        if (!path) continue;
        const pib = await getBlip(path);
        if (!pib) continue;
        const pos = (el) => el ? {
          c: +(kid(el, 'col') ? kid(el, 'col').textContent : 0), r: +(kid(el, 'row') ? kid(el, 'row').textContent : 0),
          x: +(kid(el, 'colOff') ? kid(el, 'colOff').textContent : 0) / EMU, y: +(kid(el, 'rowOff') ? kid(el, 'rowOff').textContent : 0) / EMU
        } : null;
        const shape = { pib, from: pos(kid(anchor, 'from')), to: pos(kid(anchor, 'to')), w: 0, h: 0 };
        const ext = kid(anchor, 'ext');
        if (ext) { shape.w = +ext.getAttribute('cx') / EMU; shape.h = +ext.getAttribute('cy') / EMU; }
        if (!shape.from) continue;
        out.shapes.push(shape);
      }
    }
  }
  return { styles, blips, sheets, defaultXf: 0 };
}

module.exports = { parseXlsx };
