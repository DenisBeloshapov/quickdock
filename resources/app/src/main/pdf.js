const fs = require('fs');
const path = require('path');
const os = require('os');
const { exec } = require('child_process');
const { BrowserWindow, nativeImage } = require('electron');
const { PDFDocument } = require('pdf-lib');
const XLSX = require('xlsx');
const { parseXls } = require('./xlsParse');
const { parseXlsx } = require('./xlsxParse');

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
// Если Excel не уложился в 20 с (тяжёлый файл), не ждём: убиваем наш экземпляр и сразу берём быстрый запасной рендер.
const EXCEL_TIMEOUT_MS = 20000;
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
  # PID именно нашего экземпляра Excel — чтобы при таймауте убить только его, а не Excel пользователя.
  try {
    Add-Type -Namespace QD -Name Win -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(System.IntPtr h, out int pid);'
    $xlPid = 0
    [void][QD.Win]::GetWindowThreadProcessId([System.IntPtr]$excel.Hwnd, [ref]$xlPid)
    Write-Output ('PID:' + $xlPid)
  } catch {}
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.AskToUpdateLinks = $false
  $excel.AlertBeforeOverwriting = $false
  # IgnoreRemoteRequests = игнорировать DDE-запросы от других процессов
  try { $excel.IgnoreRemoteRequests = $true } catch {}
  # AutomationSecurity = 3 (msoAutomationSecurityForceDisable) — отключаем макросы при открытии
  try { $excel.AutomationSecurity = 3 } catch {}
  try { $excel.EnableEvents = $false } catch {}
  # Пустая книга нужна, чтобы выставить ручной пересчёт ДО открытия тяжёлого файла
  # (без открытой книги Calculation задать нельзя, а открытие с автопересчётом — главная причина долгой загрузки).
  try { $blank = $excel.Workbooks.Add(); $excel.Calculation = -4135 } catch {}
  Write-Output 'STEP2:opening'
  # Workbooks.Open: Filename, UpdateLinks=0, ReadOnly=true, Format, Password, WriteResPassword,
  # IgnoreReadOnlyRecommended=true, Origin, Delimiter, Editable, Notify, Converter, AddToMru
  $wb = $excel.Workbooks.Open('${excelPathPS}', 0, $true, $null, $null, $null, $true)
  # Calculation = -4135 (xlCalculationManual) — НЕ пересчитываем формулы при открытии.
  # Это критично для тяжёлых .xls с формулами: Excel может зависнуть на минуты.
  try { $excel.Calculation = -4135 } catch {}
  try { if ($blank) { $blank.Close($false) } } catch {}
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
    exec(`powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand ${encoded}`, { timeout: EXCEL_TIMEOUT_MS, windowsHide: true }, (err, stdout, stderr) => {
      const out = (stdout || '').trim();
      // При таймауте powershell убит, а наш Excel остаётся в фоне — добиваем его по PID.
      if (err && err.killed) {
        const m = out.match(/PID:(\d+)/);
        if (m && Number(m[1]) > 0) { try { exec(`taskkill /PID ${m[1]} /F`, { windowsHide: true }, () => {}); } catch {} }
      }
      const lastLine = out.split(/\r?\n/).pop() || '';
      if (err || !lastLine.startsWith('OK')) {
        const errMsg = lastLine.startsWith('ERR:') ? lastLine.slice(4) : (err ? (err.killed ? `Excel timeout (${EXCEL_TIMEOUT_MS / 1000}s)` : err.message) : 'Unknown error');
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
 * HTML-рендер таблицы (запасной путь, когда Excel/LibreOffice недоступны или слишком медленные).
 * Значения ячеек — из библиотеки xlsx; оформление (шрифты, границы, заливки, высоты строк, картинки)
 * для .xls — из собственного разбора BIFF8 (xlsParse.js). Область печати и скрытые листы/строки/столбцы
 * учитываются так же, как в Excel. Возвращает { html, width, zoom }.
 */
async function loadStyleData(fp) {
  try {
    if (/\.xls$/i.test(fp)) return parseXls(fp);
    if (/\.xls[xm]$/i.test(fp)) return await parseXlsx(fp);
  } catch (e) { console.error('style parse failed:', e.message); }
  return null;
}

function buildXlsxHtml(fp, extra) {
  const wb = XLSX.readFile(fp, { cellStyles: true, cellDates: true });
  const sheetMeta = (wb.Workbook && wb.Workbook.Sheets) || [];
  const wbNames = (wb.Workbook && wb.Workbook.Names) || [];
  const PX_PER_CH = 7; // ширина символа «0» шрифта по умолчанию (Arial 10 / Calibri 11)
  const parts = [];
  let maxW = 0;
  let totalH = 0;

  const visible = wb.SheetNames.filter((n, i) => !(sheetMeta[i] && sheetMeta[i].Hidden));

  for (let si = 0; si < wb.SheetNames.length; si++) {
    const sheetName = wb.SheetNames[si];
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) continue;
    if (sheetMeta[si] && sheetMeta[si].Hidden) continue;

    let range = XLSX.utils.decode_range(ws['!ref']);
    const pa = wbNames.find(n => n.Name === '_xlnm.Print_Area' && n.Sheet === si && n.Ref && !/#REF/.test(n.Ref));
    if (pa) {
      try {
        const r = XLSX.utils.decode_range(pa.Ref.split('!').pop().replace(/\$/g, '').split(',')[0]);
        range = { s: { r: r.s.r, c: r.s.c }, e: { r: r.e.r, c: r.e.c } };
      } catch {}
    }
    const minRow = range.s.r, maxRow = range.e.r, minCol = range.s.c, maxCol = range.e.c;
    const ex = extra && extra.sheets[si] ? extra.sheets[si] : null;

    // --- Размеры ---
    const cols = ws['!cols'] || [];
    const colPx = (c) => {
      if (ex && ex.cols && ex.cols.has(c)) { // .xlsx: ширины из самого файла
        const xc = ex.cols.get(c);
        if (xc.hidden) return 0;
        if (xc.w) return Math.round(xc.w * PX_PER_CH);
        return 64;
      }
      const ci = cols[c];
      if (ci && ci.hidden) return 0;
      if (ci && ci.width) return Math.round(ci.width * PX_PER_CH);
      return 64;
    };
    const defH = (ex && ex.defaultH) || 15;
    const rowInfo = (r) => {
      const sr = (ws['!rows'] || [])[r];
      const er = ex ? ex.rows.get(r) : null;
      const hidden = !!((er && er.hidden) || (sr && sr.hidden));
      let pt = (er && er.h) || (sr && sr.hpt) || defH;
      return { hidden, pt };
    };
    const rowPx = (r) => { const i = rowInfo(r); return i.hidden ? 0 : i.pt * 96 / 72; };
    const xOf = (c) => { let x = 0; for (let k = 0; k < c; k++) x += colPx(k); return x; };
    const yOf = (r) => { let y = 0; for (let k = 0; k < r; k++) y += rowPx(k); return y; };
    const x0 = xOf(minCol), y0 = yOf(minRow);
    let W = 0; for (let c = minCol; c <= maxCol; c++) W += colPx(c);
    let H = 0; for (let r = minRow; r <= maxRow; r++) H += rowPx(r);

    // --- Объединённые ячейки ---
    const mergeMap = {}, covered = {};
    for (const m of ws['!merges'] || []) {
      mergeMap[`${m.s.r},${m.s.c}`] = m;
      for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) {
        if (r !== m.s.r || c !== m.s.c) covered[`${r},${c}`] = true;
      }
    }

    const styleAt = (r, c) => {
      if (!extra || !ex) return null;
      // Приоритет как в Excel: ячейка → стиль строки → стиль столбца → «Обычный».
      let xf = ex.cellXf.get(`${r},${c}`);
      if (xf == null) { const ro = ex.rows.get(r); if (ro && ro.xf != null) xf = ro.xf; }
      if (xf == null) xf = ex.colXf.get(c);
      if (xf == null) xf = extra.defaultXf != null ? extra.defaultXf : 15;
      return extra.styles[xf] || extra.styles[0] || null;
    };
    const hasValue = (r, c) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      return !!cell && cell.v !== undefined && cell.v !== null && String(cell.v) !== '';
    };

    let html = `<table style="border-collapse:collapse;table-layout:fixed;width:${W}px;"><colgroup>`;
    for (let c = minCol; c <= maxCol; c++) { const w = colPx(c); if (w > 0) html += `<col style="width:${w}px;">`; }
    html += '</colgroup>';

    for (let r = minRow; r <= maxRow; r++) {
      const ri = rowInfo(r);
      if (ri.hidden) continue;
      html += `<tr style="height:${ri.pt}pt;">`;
      for (let c = minCol; c <= maxCol; c++) {
        const key = `${r},${c}`;
        if (covered[key]) continue;
        if (colPx(c) === 0) continue;

        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        const st = styleAt(r, c);
        const m = mergeMap[key];

        // Текст
        let text = '';
        if (cell && cell.v !== undefined && cell.v !== null) {
          text = cell.w !== undefined ? String(cell.w) : String(cell.v);
          // Встроенный формат даты «m/d/yy» Excel показывает по региональным настройкам; библиотека даёт американский вид.
          // Для счетов и таблиц ожидаем дд/мм/гггг.
          if (cell.v instanceof Date && (!cell.z || /^(m\/d\/yy|mm-dd-yy|m\/d\/yyyy|General)$/i.test(cell.z))) {
            const d = cell.v;
            text = `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
          }
        }

        // Границы: у объединённой области правый край — от правой ячейки, нижний — от нижней.
        let bd = st ? { ...st.border } : null;
        if (st && m) {
          const sr = styleAt(r, m.e.c), sb = styleAt(m.e.r, c);
          if (sr) bd.r = sr.border.r;
          if (sb) bd.b = sb.border.b;
        }

        let css = 'overflow:hidden;padding:0 2px;line-height:1.15;';
        if (st) {
          const f = st.font;
          css += `font-family:'${String(f.name).replace(/'/g, '')}','Segoe UI',Arial,sans-serif;font-size:${f.size}pt;color:${f.color};`;
          if (f.bold) css += 'font-weight:bold;';
          if (f.italic) css += 'font-style:italic;';
          if (f.underline || f.strike) css += `text-decoration:${f.underline ? 'underline ' : ''}${f.strike ? 'line-through' : ''};`;
          if (st.fill) css += `background:${st.fill};`;
          if (bd.l) css += `border-left:${bd.l};`;
          if (bd.r) css += `border-right:${bd.r};`;
          if (bd.t) css += `border-top:${bd.t};`;
          if (bd.b) css += `border-bottom:${bd.b};`;
          const ha = st.halign;
          let align = (ha === 2 || ha === 6 || ha === 7) ? 'center' : ha === 3 ? 'right' : ha === 1 || ha === 4 || ha === 5 ? 'left' : null;
          if (!align) align = cell && cell.t === 'n' ? 'right' : cell && (cell.t === 'b' || cell.t === 'e') ? 'center' : 'left';
          css += `text-align:${align};`;
          css += `vertical-align:${st.valign === 0 || st.valign === 3 ? 'top' : st.valign === 1 || st.valign === 4 ? 'middle' : 'bottom'};`;
          css += st.wrap ? 'white-space:pre-wrap;word-wrap:break-word;' : 'white-space:pre;';
          if (st.indent && align === 'left') css += `padding-left:${2 + st.indent * 9}px;`;
        } else {
          // Нет данных об оформлении (например .xlsx) — тонкая сетка, как раньше.
          css += 'border:1px solid #bbb;font-family:"Segoe UI",Calibri,Arial,sans-serif;font-size:10pt;vertical-align:bottom;white-space:pre;';
          css += cell && cell.t === 'n' ? 'text-align:right;' : 'text-align:left;';
          if (cell && cell.s && cell.s.fgColor && cell.s.fgColor.rgb && cell.s.patternType === 'solid') {
            const rgb = String(cell.s.fgColor.rgb); css += `background:#${rgb.slice(-6)};`;
          }
        }

        let rs = m ? m.e.r - m.s.r + 1 : 1;
        let cs = m ? m.e.c - m.s.c + 1 : 1;

        // Перетекание длинного текста вправо в пустые соседние ячейки (как в Excel).
        if (!m && st && !st.wrap && text && cell && cell.t === 's' && (st.halign === 0 || st.halign === 1)) {
          const textW = text.length * st.font.size * 1.333 * (st.font.bold ? 0.6 : 0.54) + 6;
          let have = colPx(c);
          let k = c + 1;
          let endStyle = null;
          while (have < textW && k <= maxCol) {
            const nst = styleAt(r, k);
            if (covered[`${r},${k}`] || mergeMap[`${r},${k}`] || hasValue(r, k)) break;
            // Как в Excel: текст перетекает и через ячейки с рамкой, но не через заливку.
            if (nst && (nst.fill || (nst.border.t && !st.border.t) || (nst.border.b && !st.border.b))) break;
            have += colPx(k); cs++; k++; endStyle = nst;
          }
          // Правая граница объединённого отрезка — от последней ячейки.
          if (cs > 1 && endStyle && endStyle.border.r) css += `border-right:${endStyle.border.r};`;
        }

        html += `<td${rs > 1 ? ` rowspan="${rs}"` : ''}${cs > 1 ? ` colspan="${cs}"` : ''} style="${css}">${escapeHtml(text)}</td>`;
        c += (!m ? cs - 1 : 0);
      }
      html += '</tr>';
    }
    html += '</table>';

    // --- Картинки (логотипы, печати) ---
    let imgs = '';
    if (extra && ex) {
      for (const sh of ex.shapes) {
        const img = extra.blips[sh.pib - 1];
        if (!img) continue;
        let left, top, right, bottom;
        if (sh.anchor) { // .xls: смещения в долях ячейки
          const a = sh.anchor;
          left = xOf(a.c1) + (a.dx1 / 1024) * (colPx(a.c1) || 64) - x0;
          top = yOf(a.r1) + (a.dy1 / 256) * (rowPx(a.r1) || 20) - y0;
          right = xOf(a.c2) + (a.dx2 / 1024) * (colPx(a.c2) || 64) - x0;
          bottom = yOf(a.r2) + (a.dy2 / 256) * (rowPx(a.r2) || 20) - y0;
        } else { // .xlsx: смещения в пикселях, размер либо «до ячейки», либо абсолютный
          const f = sh.from;
          left = xOf(f.c) + f.x - x0;
          top = yOf(f.r) + f.y - y0;
          right = sh.to ? xOf(sh.to.c) + sh.to.x - x0 : left + sh.w;
          bottom = sh.to ? yOf(sh.to.r) + sh.to.y - y0 : top + sh.h;
        }
        if (right <= left || bottom <= top) continue;
        if (right < 0 || bottom < 0 || left > W || top > H) continue; // вне области печати
        imgs += `<img src="data:${img.mime};base64,${Buffer.from(img.data).toString('base64')}" style="position:absolute;left:${left.toFixed(1)}px;top:${top.toFixed(1)}px;width:${(right - left).toFixed(1)}px;height:${(bottom - top).toFixed(1)}px;z-index:2;">`;
      }
    }

    const title = visible.length > 1 ? `<div style="font:bold 14px 'Segoe UI',Arial,sans-serif;margin:0 0 6px;">${escapeHtml(sheetName)}</div>` : '';
    parts.push(`<div style="margin-bottom:16px;">${title}<div style="position:relative;width:${W}px;overflow:hidden;">${html}${imgs}</div></div>`);
    maxW = Math.max(maxW, W);
    totalH += H + 16 + (title ? 24 : 0);
  }

  // Чёткость: удваиваем масштаб, пока картинка укладывается в лимит высоты скриншота.
  const zoom = totalH * 2 < 15000 ? 2 : 1;
  const html = `<!doctype html><html><head><meta charset="utf-8"><style>
    *{box-sizing:border-box;}
    html,body{margin:0;padding:0;background:#fff;overflow:hidden;}
    body{padding:8px;zoom:${zoom};color:#000;}
    ::-webkit-scrollbar{display:none;}
    td{margin:0;}
    img{image-rendering:auto;}
  </style></head><body>${parts.join('')}</body></html>`;
  return { html, width: Math.ceil((maxW + 16) * zoom), zoom };
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
  const extra = await loadStyleData(fp);
  const built = buildXlsxHtml(fp, extra);
  const { png } = await renderHtmlToPng(built.html, { width: built.width });
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
