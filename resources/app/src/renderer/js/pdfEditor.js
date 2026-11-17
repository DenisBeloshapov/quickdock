(function () {
  const { qdPdfEditor } = window;

  let currentPath = '';
  let pdfBytes = null;
  let pdfDoc = null;
  let pages = [];
  let initialOrder = [];

  function getPdfJs() {
    if (!window.pdfjsLib) throw new Error('pdf.js не загружен');
    if (!window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = '../../build/pdfjs/pdf.worker.min.js';
    }
    return window.pdfjsLib;
  }

  async function loadFile(filePath) {
    currentPath = filePath;
    document.getElementById('fileName').textContent = filePath.split(/[\\/]/).pop();
    setStatus('Загрузка PDF…', '');
    const grid = document.getElementById('pagesGrid');
    grid.innerHTML = '<div class="loading">Загрузка…</div>';

    const data = await qdPdfEditor.getData(filePath);
    if (!data.ok) {
      grid.innerHTML = '<div class="loading">Ошибка: ' + escapeHtml(data.error) + '</div>';
      setStatus('Ошибка: ' + data.error, 'error');
      return;
    }
    const bytesRes = await qdPdfEditor.getBytes(filePath);
    if (!bytesRes.ok) {
      grid.innerHTML = '<div class="loading">Ошибка: ' + escapeHtml(bytesRes.error) + '</div>';
      return;
    }
    pdfBytes = new Uint8Array(bytesRes.bytes);

    try {
      const lib = getPdfJs();
      const buf = pdfBytes.slice(0);
      pdfDoc = await lib.getDocument({ data: buf }).promise;
    } catch (e) {
      grid.innerHTML = '<div class="loading">Ошибка: ' + escapeHtml(e.message) + '</div>';
      setStatus('Ошибка: ' + e.message, 'error');
      return;
    }

    pages = [];
    for (let i = 0; i < pdfDoc.numPages; i++) {
      pages.push({ oldIndex: i, rotation: 0, deleted: false });
    }
    initialOrder = pages.map((p) => ({ ...p }));
    document.getElementById('infoPages').textContent = 'Страниц: ' + pdfDoc.numPages;
    updateChangesCount();
    await renderPages();
    setStatus('Загружено страниц: ' + pdfDoc.numPages, 'success');
  }

  async function renderPages() {
    const grid = document.getElementById('pagesGrid');
    grid.innerHTML = '';
    const visible = pages.filter((p) => !p.deleted);
    if (!visible.length) {
      grid.innerHTML = '<div class="loading">Все страницы удалены</div>';
      return;
    }
    for (let pos = 0; pos < visible.length; pos++) {
      const card = await createPageCard(visible[pos], pos);
      grid.appendChild(card);
    }
  }

  async function createPageCard(page, pos) {
    const lib = getPdfJs();
    const pdfPage = await pdfDoc.getPage(page.oldIndex + 1);
    const baseVp = pdfPage.getViewport({ scale: 1 });
    const targetH = 240;
    const scale = targetH / baseVp.height;
    const vp = pdfPage.getViewport({ scale, rotation: page.rotation });
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    canvas.width = vp.width; canvas.height = vp.height;
    await pdfPage.render({ canvasContext: ctx, viewport: vp }).promise;

    const card = document.createElement('div');
    card.className = 'page-card';
    card.draggable = true;
    card.dataset.pos = String(pos);
    card.innerHTML =
      '<div class="page-head"><span class="page-num">#' + (pos + 1) + '</span>' +
      '<button class="page-rotate" title="Повернуть на 90° по часовой">↻</button></div>' +
      '<div class="page-canvas-wrap"></div>' +
      '<div class="page-foot"><span>исх. стр. ' + (page.oldIndex + 1) + '</span>' +
      '<span class="page-rot-badge">' + (page.rotation ? page.rotation + '°' : '') + '</span>' +
      '<button class="page-del" title="Удалить страницу">×</button></div>';
    card.querySelector('.page-canvas-wrap').appendChild(canvas);

    card.querySelector('.page-rotate').addEventListener('click', (e) => {
      e.stopPropagation();
      page.rotation = (page.rotation + 90) % 360;
      updateChangesCount();
      rerenderCard(card, page, pos);
    });
    card.querySelector('.page-del').addEventListener('click', (e) => {
      e.stopPropagation();
      if (!qdPdfEditor.confirm('Удалить страницу ' + (pos + 1) + '?')) return;
      page.deleted = true;
      updateChangesCount();
      renderPages();
    });
    card.addEventListener('click', () => {
      document.querySelectorAll('.page-card').forEach((c) => c.classList.remove('selected'));
      card.classList.add('selected');
    });
    setupDragHandlers(card, pos);
    return card;
  }

  async function rerenderCard(card, page, pos) {
    const wrap = card.querySelector('.page-canvas-wrap');
    wrap.innerHTML = '';
    const lib = getPdfJs();
    const pdfPage = await pdfDoc.getPage(page.oldIndex + 1);
    const baseVp = pdfPage.getViewport({ scale: 1 });
    const targetH = 240;
    const scale = targetH / baseVp.height;
    const vp = pdfPage.getViewport({ scale, rotation: page.rotation });
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    canvas.width = vp.width; canvas.height = vp.height;
    await pdfPage.render({ canvasContext: ctx, viewport: vp }).promise;
    wrap.appendChild(canvas);
    card.querySelector('.page-rot-badge').textContent = page.rotation ? page.rotation + '°' : '';
    card.querySelector('.page-num').textContent = '#' + (pos + 1);
  }

  let dragSrcPos = -1;
  function setupDragHandlers(card, pos) {
    card.addEventListener('dragstart', (e) => {
      dragSrcPos = pos;
      card.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', String(pos));
    });
    card.addEventListener('dragend', () => {
      card.classList.remove('dragging');
      document.querySelectorAll('.page-card').forEach((c) => c.classList.remove('drag-over'));
    });
    card.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      if (dragSrcPos !== pos) card.classList.add('drag-over');
    });
    card.addEventListener('dragleave', () => card.classList.remove('drag-over'));
    card.addEventListener('drop', (e) => {
      e.preventDefault();
      card.classList.remove('drag-over');
      if (dragSrcPos < 0 || dragSrcPos === pos) return;
      const visible = pages.filter((p) => !p.deleted);
      const moved = visible[dragSrcPos];
      visible.splice(dragSrcPos, 1);
      visible.splice(pos, 0, moved);
      const deleted = pages.filter((p) => p.deleted);
      pages = visible.concat(deleted);
      dragSrcPos = -1;
      updateChangesCount();
      renderPages();
    });
  }

  function updateChangesCount() {
    let changes = 0;
    const visible = pages.filter((p) => !p.deleted);
    for (let i = 0; i < visible.length; i++) {
      if (visible[i].oldIndex !== i) { changes++; break; }
    }
    for (const p of visible) { if (p.rotation % 360 !== 0) changes++; }
    changes += pages.filter((p) => p.deleted).length;
    document.getElementById('infoChanges').textContent = 'Изменения: ' + changes;
  }

  function setStatus(msg, kind) {
    const el = document.getElementById('status');
    el.textContent = msg;
    el.className = 'status mono' + (kind ? ' ' + kind : '');
  }

  function buildPayload() {
    const visible = pages.filter((p) => !p.deleted);
    const order = visible.map((p) => p.oldIndex);
    const rotations = {};
    for (const p of visible) {
      if (p.rotation % 360 !== 0) rotations[p.oldIndex] = p.rotation;
    }
    return { path: currentPath, order, rotations };
  }

  document.getElementById('btnSave').addEventListener('click', async () => {
    setStatus('Сохранение…', '');
    const payload = buildPayload();
    payload.saveAs = true;
    const r = await qdPdfEditor.save(payload);
    if (r.ok) {
      setStatus('Сохранено: ' + r.path, 'success');
      document.getElementById('fileName').textContent = r.path.split(/[\\/]/).pop();
      currentPath = r.path;
    } else if (r.canceled) {
      setStatus('Сохранение отменено', 'warning');
    } else {
      setStatus('Ошибка: ' + r.error, 'error');
    }
  });

  document.getElementById('btnOverwrite').addEventListener('click', async () => {
    if (!qdPdfEditor.confirm('Перезаписать оригинальный файл? Это действие необратимо.')) return;
    setStatus('Перезапись…', '');
    const r = await qdPdfEditor.overwrite(buildPayload());
    if (r.ok) {
      setStatus('Файл перезаписан: ' + r.path, 'success');
      await loadFile(r.path);
    } else {
      setStatus('Ошибка: ' + r.error, 'error');
    }
  });

  document.getElementById('btnReset').addEventListener('click', () => {
    if (!qdPdfEditor.confirm('Сбросить все изменения?')) return;
    pages = initialOrder.map((p) => ({ ...p }));
    updateChangesCount();
    renderPages();
    setStatus('Изменения сброшены', 'warning');
  });

  function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  qdPdfEditor.onLoaded((p) => { if (p && p.path) loadFile(p.path); });
})();
