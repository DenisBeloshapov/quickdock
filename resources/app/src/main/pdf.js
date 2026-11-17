const fs = require('fs');
const path = require('path');
const os = require('os');
const { exec } = require('child_process');
const { BrowserWindow, nativeImage } = require('electron');
const { PDFDocument } = require('pdf-lib');
const XLSX = require('xlsx');

const A4_W = 595.28;
const A4_H = 841.89;
const A4_MARGIN = 24;

function classify(fp) {
  const ext = path.extname(fp).toLowerCase();
  if (['.png','.jpg','.jpeg','.jfif','.bmp','.gif','.webp'].includes(ext)) return 'image';
  if (ext === '.docx') return 'docx';
  if (ext === '.doc') return 'doc';
  if (ext === '.xlsx' || ext === '.xls') return 'xlsx';
  if (ext === '.pdf') return 'pdf';
  return 'unsupported';
}

// Флаги: COM-объект Word/Excel недоступен в системе. После первой постоянной ошибки
// (COM не зарегистрирован = Office не установлен) больше не пытаемся — экономит ~30с на каждом файле.
let wordUnavailable = false;
let excelUnavailable = false;

// === Поиск LibreOffice (кросс-платформенный fallback) ===
// LibreOffice даёт близкое к оригиналу соответствие для .doc/.docx/.xls/.xlsx.
// На Windows ищем в Program Files, в PATH, через реестр.
let libreofficePath = null;
let libreofficeChecked = false;

function findLibreOffice() {
  if (libreofficeChecked) return libreofficePath;
  libreofficeChecked = true;
  // 1. soffice в PATH
  try {
    const r = require('child_process').execSync(
      process.platform === 'win32' ? 'where soffice.exe' : 'which soffice',
      { stdio: ['ignore', 'pipe', 'ignore'] }
    ).toString().trim();
    if (r) { libreofficePath = r.split(/\r?\n/)[0]; return libreofficePath; }
  } catch {}
  if (process.platform === 'win32') {
    // 2. Стандартные locations Program Files
    const candidates = [
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ];
    for (const c of candidates) {
      try { if (fs.existsSync(c)) { libreofficePath = c; return libreofficePath; } } catch {}
    }
    // 3. Через реестр
    try {
      const out = require('child_process').execSync(
        'powershell.exe -NoProfile -Command "(Get-ItemProperty \'HKLM:\\SOFTWARE\\LibreOffice\\UNO\\InstallPath\' -ErrorAction Stop).\'(default)\'"',
        { stdio: ['ignore', 'pipe', 'ignore'] }
      ).toString().trim();
      if (out) {
        const cand = out + '\\soffice.exe';
        try { if (fs.existsSync(cand)) { libreofficePath = cand; return libreofficePath; } } catch {}
      }
    } catch {}
  }
  return libreofficePath;
}

// === Конвертация через LibreOffice (soffice --headless --convert-to pdf) ===
// Кросс-платформенный fallback, когда Word/Excel COM недоступен.
function convertViaLibreOfficeToPdf(srcPath) {
  return new Promise((resolve, reject) => {
    const loBin = findLibreOffice();
    if (!loBin) { reject(new Error('LibreOffice not found')); return; }
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'qd-lo-'));
    const baseName = path.basename(srcPath).replace(/\.[^.]+$/, '');
    const expectedPdf = path.join(tmpDir, baseName + '.pdf');
    // Кавычки: на Windows используем двойные, на Unix — одинарные с экранированием
    const q = (p) => process.platform === 'win32' ? '"' + p + '"' : "'" + p.replace(/'/g, "'\\''") + "'";
    const cmd = q(loBin) + ' --headless --norestore --nofirststartwizard --convert-to pdf --outdir ' + q(tmpDir) + ' ' + q(srcPath);
    exec(cmd, { timeout: 90000, windowsHide: true, maxBuffer: 4 * 1024 * 1024 }, (err) => {
      if (err || !fs.existsSync(expectedPdf)) {
        try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
        reject(new Error('LibreOffice convert failed: ' + (err ? (err.killed ? 'timeout' : err.message) : 'no output')));
        return;
      }
      resolve(expectedPdf);
    });
  });
}

function convertViaWordToPdf(wordPath) {
  if (wordUnavailable) return Promise.reject(new Error('Word unavailable (cached)'));
  return new Promise((resolve, reject) => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'qd-word-'));
    const baseName = path.basename(wordPath).replace(/\.(docx?)$/i, '');
    const pdfPath = path.join(tmpDir, baseName + '.pdf');
    // Одинарные кавычки (literal strings) — экранируем только ' как ''. Не трогаем слэши.
    const wordPathPS = wordPath.replace(/'/g, "''");
    const pdfPathPS = pdfPath.replace(/'/g, "''");
    const psScript = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$word = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open('${wordPathPS}', $false, $true)
  $doc.SaveAs2('${pdfPathPS}', 17)
  $doc.Close($false)
  $word.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
  Write-Output 'OK'
} catch {
  try { if ($word) { $word.Quit() } } catch {}
  Write-Output ('ERR:' + $_.Exception.Message)
  exit 1
}`;
    // -EncodedCommand (base64-UTF-16LE) — корректная обработка кириллицы в путях.
    const encoded = Buffer.from(psScript, 'utf16le').toString('base64');
    exec(`powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand ${encoded}`, { timeout: 90000, windowsHide: true }, (err, stdout) => {
      const out = (stdout || '').trim();
      const lastLine = out.split(/\r?\n/).pop() || '';
      if (err || !lastLine.startsWith('OK')) {
        const errMsg = lastLine.startsWith('ERR:') ? lastLine.slice(4) : (err ? (err.killed ? 'Word timeout (90s)' : err.message) : 'Unknown error');
        try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
        if (/COM|80080005|80040154|Microsoft Word|not registered/i.test(errMsg)) wordUnavailable = true;
        reject(new Error('Word PDF export failed: ' + errMsg + '. Is Microsoft Word installed?'));
        return;
      }
      if (!fs.existsSync(pdfPath)) { try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {} reject(new Error('Word PDF export produced no output file')); return; }
      resolve(pdfPath);
    });
  });
}

// === Конвертация .xls/.xlsx через Microsoft Excel COM ===
// Даёт попиксельное соответствие тому, что видит пользователь в Excel.
// Использует максимально надёжные настройки, чтобы избежать зависаний на
// нестандартных .xls (внешние ссылки, макросы, DDE-запросы, формулы).
function convertViaExcelToPdf(excelPath) {
  if (excelUnavailable) return Promise.reject(new Error('Excel unavailable (cached)'));
  return new Promise((resolve, reject) => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'qd-excel-'));
    const baseName = path.basename(excelPath).replace(/\.(xlsx?|xlsm)$/i, '');
    const pdfPath = path.join(tmpDir, baseName + '.pdf');
    const excelPathPS = excelPath.replace(/'/g, "''");
    const pdfPathPS = pdfPath.replace(/'/g, "''");
    const psScript = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$excel = $null
$wb = $null
try {
  Write-Output 'STEP1:creating'
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.AskToUpdateLinks = $false
  $excel.AlertBeforeOverwriting = $false
  # IgnoreRemoteRequests = игнорировать DDE-запросы от других процессов
  try { $excel.IgnoreRemoteRequests = $true } catch {}
  # AutomationSecurity = 3 (msoAutomationSecurityForceDisable) — отключаем макросы при открытии
  try { $excel.AutomationSecurity = 3 } catch {}
  Write-Output 'STEP2:opening'
  # Workbooks.Open: Filename, UpdateLinks=0, ReadOnly=true, Format, Password, WriteResPassword,
  # IgnoreReadOnlyRecommended=true, Origin, Delimiter, Editable, Notify, Converter, AddToMru
  $wb = $excel.Workbooks.Open('${excelPathPS}', 0, $true, $null, $null, $null, $true)
  # Calculation = -4135 (xlCalculationManual) — НЕ пересчитываем формулы при открытии.
  # Это критично для тяжёлых .xls с формулами: Excel может зависнуть на минуты.
  try { $excel.Calculation = -4135 } catch {}
  Write-Output 'STEP3:exporting'
  # ExportAsFixedFormat: Type=0 (xlTypePDF), Quality=0 (xlQualityStandard),
  # IncludeDocProperties=true, IgnorePrintAreas=false, From, To, OpenAfterPublish=false
  $wb.ExportAsFixedFormat(0, '${pdfPathPS}', 0, $false, $false)
  Write-Output 'STEP4:closing'
  $wb.Close($false)
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  Write-Output 'OK'
} catch {
  Write-Output ('ERR:' + $_.Exception.Message)
  try { if ($wb) { $wb.Close($false) } } catch {}
  try { if ($excel) { $excel.Quit() } } catch {}
  exit 1
}`;
    const encoded = Buffer.from(psScript, 'utf16le').toString('base64');
    exec(`powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand ${encoded}`, { timeout: 30000, windowsHide: true }, (err, stdout, stderr) => {
      const out = (stdout || '').trim();
      const lastLine = out.split(/\r?\n/).pop() || '';
      if (err || !lastLine.startsWith('OK')) {
        const errMsg = lastLine.startsWith('ERR:') ? lastLine.slice(4) : (err ? (err.killed ? 'Excel timeout (30s)' : err.message) : 'Unknown error');
        try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
        // Кэшируем недоступность только для постоянных ошибок: COM-объект не зарегистрирован.
        // Таймаут НЕ кэшируем — может быть временной проблемой (занятый файл).
        if (/COM|80080005|80040154|Microsoft Excel|not registered/i.test(errMsg)) {
          excelUnavailable = true;
        }
        reject(new Error('Excel PDF export failed: ' + errMsg));
        return;
      }
      if (!fs.existsSync(pdfPath)) { try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {} reject(new Error('Excel PDF export produced no output file')); return; }
      resolve(pdfPath);
    });
  });
}

function renderHtmlToPng(html, opts = {}) {
  const width = opts.width || 1240;
  return new Promise((resolve, reject) => {
    const win = new BrowserWindow({
      width, height: 800, show: false,
      skipTaskbar: true, focusable: false,
      webPreferences: { offscreen: true, sandbox: false, contextIsolation: false }
    });
    const timeout = setTimeout(() => { try { win.destroy(); } catch {}; reject(new Error('Render timeout')); }, 30000);
    win.webContents.once('did-finish-load', async () => {
      try {
        await new Promise(r => setTimeout(r, 200));
        const h = await win.webContents.executeJavaScript(`Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)`);
        const fh = Math.min(Math.max(h, 200), 16384);
        win.setSize(width, fh);
        await new Promise(r => setTimeout(r, 300));
        const img = await win.webContents.capturePage();
        clearTimeout(timeout);
        const png = img.toPNG();
        win.destroy();
        resolve({ png, width, height: fh });
      } catch (e) { clearTimeout(timeout); try { win.destroy(); } catch {}; reject(e); }
    });
    win.webContents.once('did-fail-load', (_e, c, d) => { clearTimeout(timeout); try { win.destroy(); } catch {}; reject(new Error(`Load failed: ${c} ${d||''}`)); });
    win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
  });
}

/**
 * Build an HTML table from an XLSX file — manually, NOT using sheet_to_html.
 * Only renders the ACTIVE range (min_row..max_row, min_col..max_col) — cells
 * outside that range (empty trailing rows/cols) are excluded.
 * Handles merged cells (rowspan/colspan), dates, numbers, text.
 */
function buildXlsxHtml(fp) {
  const wb = XLSX.readFile(fp, { cellStyles: true, cellDates: true });
  const parts = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    // Get the active range — only cells with data
    const ref = ws['!ref'] || '';
    if (!ref) continue;
    const range = XLSX.utils.decode_range(ref);
    const minRow = range.s.r;
    const maxRow = range.e.r;
    const minCol = range.s.c;
    const maxCol = range.e.c;

    // Build merged cell lookup
    const merges = ws['!merges'] || [];
    const mergeMap = {};
    const mergedCells = {};
    for (const m of merges) {
      const key = `${m.s.r},${m.s.c}`;
      mergeMap[key] = { rowspan: m.e.r - m.s.r + 1, colspan: m.e.c - m.s.c + 1 };
      for (let r = m.s.r; r <= m.e.r; r++) {
        for (let c = m.s.c; c <= m.e.c; c++) {
          if (r === m.s.r && c === m.s.c) continue;
          mergedCells[`${r},${c}`] = true;
        }
      }
    }

    // Column widths
    const colWidths = ws['!cols'] || [];

    let html = `<table style="border-collapse:collapse;width:100%;font-size:11px;">`;

    // Colgroup for column widths
    html += '<colgroup>';
    for (let c = minCol; c <= maxCol; c++) {
      const w = colWidths[c] ? Math.round((colWidths[c].wpx || 80)) : 80;
      html += `<col style="width:${w}px;">`;
    }
    html += '</colgroup>';

    for (let r = minRow; r <= maxRow; r++) {
      html += '<tr>';
      for (let c = minCol; c <= maxCol; c++) {
        const cellKey = `${r},${c}`;
        if (mergedCells[cellKey]) continue; // skip cells covered by merge

        const cellAddr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[cellAddr];
        let value = '';
        let style = 'padding:4px 6px;border:1px solid #aaa;vertical-align:top;';

        if (cell) {
          if (cell.v instanceof Date) {
            const d = cell.v;
            const dd = String(d.getDate()).padStart(2, '0');
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const yy = d.getFullYear();
            value = `${dd}/${mm}/${yy}`;
          } else if (typeof cell.v === 'number') {
            value = String(cell.v);
          } else if (cell.v !== null && cell.v !== undefined) {
            value = String(cell.v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
          }
          // Check if it looks like a number with formatting
          if (cell.z && typeof cell.v === 'number') {
            try { value = XLSX.utils.format_cell(cell); } catch {}
          }
        }

        // Check for bold style
        if (cell && cell.s) {
          if (cell.s.font && cell.s.font.bold) style += 'font-weight:bold;';
          if (cell.s.alignment) {
            if (cell.s.alignment.horizontal === 'center') style += 'text-align:center;';
            if (cell.s.alignment.horizontal === 'right') style += 'text-align:right;';
          }
          if (cell.s.fill && cell.s.fill.fgColor && cell.s.fill.fgColor.rgb) {
            const bg = cell.s.fill.fgColor.rgb;
            if (bg !== '00000000' && bg.length === 8) style += `background:#${bg.slice(2)};`;
          }
        }

        const m = mergeMap[cellKey];
        const rs = m ? ` rowspan="${m.rowspan}"` : '';
        const cs = m ? ` colspan="${m.colspan}"` : '';

        html += `<td${rs}${cs} style="${style}">${value}</td>`;
      }
      html += '</tr>';
    }
    html += '</table>';

    parts.push(`<div style="margin-bottom:20px;font-family:"Segoe UI","Calibri",sans-serif;color:#1a1a1a;background:#fff;padding:20px;">
      ${wb.SheetNames.length > 1 ? `<div style="font-size:14px;font-weight:bold;margin-bottom:8px;">${escapeHtml(sheetName)}</div>` : ''}
      ${html}
    </div>`);
  }

  return `<!doctype html><html><head><meta charset="utf-8"><style>
    *{box-sizing:border-box;}
    body{font-family:"Segoe UI","Calibri",sans-serif;font-size:11px;color:#1a1a1a;background:#fff;padding:0;margin:0;}
    table{border-collapse:collapse;table-layout:fixed;}
    td{word-wrap:break-word;overflow:hidden;}
  </style></head><body>${parts.join('')}</body></html>`;
}

async function embedPngIntoA4(pdf, pngBytes) {
  const embedded = await pdf.embedPng(pngBytes);
  const imgW = embedded.width;
  const imgH = embedded.height;
  const availW = A4_W - A4_MARGIN * 2;
  const availH = A4_H - A4_MARGIN * 2;
  const scale = Math.min(availW / imgW, availH / imgH);
  const drawW = imgW * scale;
  const drawH = imgH * scale;
  const x = (A4_W - drawW) / 2;
  const y = (A4_H - drawH) / 2;
  const page = pdf.addPage([A4_W, A4_H]);
  page.drawImage(embedded, { x, y, width: drawW, height: drawH });
}

async function renderImage(pdf, fp) {
  const ni = nativeImage.createFromPath(fp);
  if (ni.isEmpty()) throw new Error('Cannot read image: ' + path.basename(fp));
  await embedPngIntoA4(pdf, ni.toPNG());
}
async function renderDocx(pdf, fp) {
  // 1. Пробуем Microsoft Word COM (Windows + MS Office).
  // 2. Если Word недоступен — LibreOffice (кросс-платформенный fallback).
  // 3. Если LibreOffice тоже нет — ошибка (mammoth не используем: он ломает layout).
  let tempPdfPath = null;
  try {
    tempPdfPath = await convertViaWordToPdf(fp);
  } catch (wordErr) {
    console.error('Word COM failed, trying LibreOffice:', wordErr.message);
    try {
      tempPdfPath = await convertViaLibreOfficeToPdf(fp);
    } catch (loErr) {
      console.error('LibreOffice failed:', loErr.message);
      throw new Error('Word и LibreOffice недоступны. Установите Microsoft Word или LibreOffice для конвертации .doc/.docx');
    }
  }
  try { await renderPdf(pdf, tempPdfPath); } finally { try { fs.rmSync(path.dirname(tempPdfPath), { recursive: true, force: true }); } catch {} }
}
async function renderDoc(pdf, fp) {
  // То же самое что renderDocx — .doc тоже идёт через Word COM / LibreOffice.
  let tempPdfPath = null;
  try {
    tempPdfPath = await convertViaWordToPdf(fp);
  } catch (wordErr) {
    console.error('Word COM failed, trying LibreOffice:', wordErr.message);
    try {
      tempPdfPath = await convertViaLibreOfficeToPdf(fp);
    } catch (loErr) {
      console.error('LibreOffice failed:', loErr.message);
      throw new Error('Word и LibreOffice недоступны. Установите Microsoft Word или LibreOffice для конвертации .doc/.docx');
    }
  }
  try { await renderPdf(pdf, tempPdfPath); } finally { try { fs.rmSync(path.dirname(tempPdfPath), { recursive: true, force: true }); } catch {} }
}
async function renderXlsx(pdf, fp) {
  // 1. Microsoft Excel COM — попиксельное соответствие Excel.
  // 2. Если недоступен — LibreOffice (кросс-платформенный fallback, тоже точное соответствие).
  // 3. Если оба недоступны — собственный HTML-рендер (только активные ячейки, без сложного layout).
  if (!excelUnavailable) {
    try {
      const tempPdfPath = await convertViaExcelToPdf(fp);
      try {
        await renderPdf(pdf, tempPdfPath);
        return; // Успех.
      } finally {
        try { fs.rmSync(path.dirname(tempPdfPath), { recursive: true, force: true }); } catch {}
      }
    } catch (excelErr) {
      console.error('Excel COM failed, trying LibreOffice:', excelErr.message);
    }
  }
  // Пробуем LibreOffice как fallback.
  try {
    const tempPdfPath = await convertViaLibreOfficeToPdf(fp);
    try {
      await renderPdf(pdf, tempPdfPath);
      return;
    } finally {
      try { fs.rmSync(path.dirname(tempPdfPath), { recursive: true, force: true }); } catch {}
    }
  } catch (loErr) {
    console.error('LibreOffice failed, falling back to HTML renderer:', loErr.message);
  }
  // Последний fallback — собственный HTML-рендер таблицы. Только активные ячейки.
  const html = buildXlsxHtml(fp);
  const { png } = await renderHtmlToPng(html, { width: 900 });
  await embedPngIntoA4(pdf, png);
}
async function renderPdf(pdf, fp) {
  const src = await PDFDocument.load(fs.readFileSync(fp), { ignoreEncryption: true });
  const idx = src.getPageIndices();
  if (idx.length === 0) return;
  const copied = await pdf.copyPages(src, idx);
  for (const p of copied) pdf.addPage(p);
}
function escapeHtml(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

async function buildPdf(filePaths, outPath) {
  const pdf = await PDFDocument.create();
  const errors = [];
  for (const fp of filePaths) {
    try {
      const kind = classify(fp);
      if (kind === 'image') await renderImage(pdf, fp);
      else if (kind === 'docx') await renderDocx(pdf, fp);
      else if (kind === 'doc') await renderDoc(pdf, fp);
      else if (kind === 'xlsx') await renderXlsx(pdf, fp);
      else if (kind === 'pdf') await renderPdf(pdf, fp);
      else throw new Error('Unsupported: ' + path.extname(fp));
    } catch (e) { errors.push(`${path.basename(fp)}: ${e.message}`); }
  }
  if (pdf.getPageCount() === 0) throw new Error('No pages. ' + errors.join('; '));
  fs.writeFileSync(outPath, await pdf.save());
  return { outPath, errors };
}
module.exports = { buildPdf, classify };
