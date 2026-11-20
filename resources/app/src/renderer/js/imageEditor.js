(function () {
  const api = window.qdImageEditor;

  const MIN_CROP = 0.03;        // минимальный размер рамки (доля стороны)
  const PREVIEW_MAX = 1600;     // длинная сторона превью (оригинал в памяти не держим)
  const FULL = () => ({ x: 0, y: 0, w: 1, h: 1 });

  /** @type {{id:number,name:string,mime:string,bytes:Uint8Array,exif:number,pv:HTMLCanvasElement,pw:number,ph:number,w:number,h:number,rot:number,crop:{x:number,y:number,w:number,h:number}}[]} */
  let items = [];
  let selectedId = null;
  let nextId = 1;
  let busy = false;

  const $ = (id) => document.getElementById(id);
  const cardsEl = $('cards'), stageCanvas = $('stageCanvas'), cropWrap = $('cropWrap'), cropEl = $('crop');

  // ---------- утилиты ----------
  function setStatus(msg, kind) {
    const el = $('status');
    el.textContent = msg;
    el.className = 'status mono' + (kind ? ' ' + kind : '');
  }
  const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
  const isFull = (c) => c.x <= 0.0005 && c.y <= 0.0005 && c.w >= 0.999 && c.h >= 0.999;
  const isChanged = (it) => it.rot !== 0 || !isFull(it.crop);
  const baseName = (n) => String(n).replace(/\.[^.]+$/, '');

  function sniffMime(b) {
    if (b.length > 3 && b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return 'image/jpeg';
    if (b.length > 4 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'image/png';
    if (b.length > 4 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'image/gif';
    if (b.length > 2 && b[0] === 0x42 && b[1] === 0x4D) return 'image/bmp';
    if (b.length > 12 && b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return 'image/webp';
    return null;
  }

  // Ориентация из EXIF (1 = обычная). Нужна, чтобы не вставлять «боком» оригинальный JPEG напрямую.
  function exifOrientation(b) {
    if (b.length < 4 || b[0] !== 0xFF || b[1] !== 0xD8) return 1;
    let i = 2;
    while (i + 4 < b.length) {
      if (b[i] !== 0xFF) break;
      const marker = b[i + 1];
      const len = (b[i + 2] << 8) | b[i + 3];
      if (marker === 0xE1 && b[i + 4] === 0x45 && b[i + 5] === 0x78 && b[i + 6] === 0x69 && b[i + 7] === 0x66) {
        const t = i + 10;
        const le = b[t] === 0x49;
        const r16 = (o) => le ? (b[o] | (b[o + 1] << 8)) : ((b[o] << 8) | b[o + 1]);
        const r32 = (o) => le ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) : ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]);
        const ifd = t + r32(t + 4);
        const n = r16(ifd);
        for (let k = 0; k < n; k++) {
          const e = ifd + 2 + k * 12;
          if (r16(e) === 0x0112) return r16(e + 8) || 1;
        }
        return 1;
      }
      if (marker === 0xDA) break;
      i += 2 + len;
    }
    return 1;
  }

  // Рисует повёрнутую (rot, по часовой) и обрезанную (crop — доли повёрнутого изображения) картинку в ctx размером ow×oh.
  function paint(ctx, src, sw, sh, rot, crop, ow, oh) {
    const swap = rot % 180 !== 0;
    const rw = swap ? sh : sw, rh = swap ? sw : sh;
    const cx = crop.x * rw, cy = crop.y * rh, cw = crop.w * rw, ch = crop.h * rh;
    ctx.save();
    ctx.scale(ow / cw, oh / ch);
    ctx.translate(-cx, -cy);
    if (rot === 90) { ctx.translate(rw, 0); ctx.rotate(Math.PI / 2); }
    else if (rot === 180) { ctx.translate(rw, rh); ctx.rotate(Math.PI); }
    else if (rot === 270) { ctx.translate(0, rh); ctx.rotate(-Math.PI / 2); }
    ctx.drawImage(src, 0, 0, sw, sh);
    ctx.restore();
  }

  const rotatedDims = (it) => (it.rot % 180 ? { w: it.h, h: it.w } : { w: it.w, h: it.h });

  // ---------- загрузка файлов ----------
  async function addBytes(name, bytes) {
    const mime = sniffMime(bytes);
    if (!mime) { setStatus('Не изображение: ' + name, 'warning'); return false; }
    let bmp;
    try {
      bmp = await createImageBitmap(new Blob([bytes], { type: mime }), { imageOrientation: 'from-image' });
    } catch (e) { setStatus('Не удалось открыть: ' + name, 'error'); return false; }
    const w = bmp.width, h = bmp.height;
    const k = Math.min(1, PREVIEW_MAX / Math.max(w, h));
    const pv = document.createElement('canvas');
    pv.width = Math.max(1, Math.round(w * k)); pv.height = Math.max(1, Math.round(h * k));
    pv.getContext('2d').drawImage(bmp, 0, 0, pv.width, pv.height);
    bmp.close();
    items.push({
      id: nextId++, name, mime, bytes, exif: mime === 'image/jpeg' ? exifOrientation(bytes) : 1,
      pv, pw: pv.width, ph: pv.height, w, h, rot: 0, crop: FULL(),
    });
    return true;
  }

  async function addPaths(paths) {
    if (!paths || !paths.length) return;
    setStatus('Загрузка…', '');
    let added = 0;
    for (const p of paths) {
      const r = await api.read(p);
      if (!r.ok) { setStatus('Ошибка: ' + r.error, 'error'); continue; }
      if (await addBytes(r.name, new Uint8Array(r.bytes))) added++;
    }
    afterAdd(added);
  }

  async function addFiles(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    setStatus('Загрузка…', '');
    let added = 0;
    for (const f of files) {
      const buf = new Uint8Array(await f.arrayBuffer());
      if (await addBytes(f.name, buf)) added++;
    }
    afterAdd(added);
  }

  function afterAdd(added) {
    if (added && selectedId === null) selectedId = items[0].id;
    else if (added) selectedId = items[items.length - added].id; // выбираем первое из добавленных
    renderAll();
    if (added) setStatus('Добавлено фото: ' + added, 'success');
  }

  // ---------- отрисовка ----------
  function renderAll() {
    renderCards();
    renderStage();
    updateInfo();
  }

  function updateInfo() {
    $('infoCount').textContent = 'Фото: ' + items.length;
    $('infoChanges').textContent = 'Изменено: ' + items.filter(isChanged).length;
    $('btnSave').disabled = busy || !items.length;
    $('btnReset').disabled = busy || !items.some(isChanged);
  }

  function drawThumb(it, canvas) {
    const rd = rotatedDims(it);
    const asp = (it.crop.w * rd.w) / (it.crop.h * rd.h);
    const maxW = 360, maxH = 280;
    let ow = maxW, oh = Math.round(maxW / asp);
    if (oh > maxH) { oh = maxH; ow = Math.round(maxH * asp); }
    canvas.width = Math.max(1, ow); canvas.height = Math.max(1, oh);
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, ow, oh);
    paint(ctx, it.pv, it.pw, it.ph, it.rot, it.crop, canvas.width, canvas.height);
  }

  function cardSizeText(it) {
    const rd = rotatedDims(it);
    const cw = Math.round(it.crop.w * rd.w), ch = Math.round(it.crop.h * rd.h);
    return cw + '×' + ch;
  }

  function renderCards() {
    cardsEl.innerHTML = '';
    items.forEach((it, pos) => {
      const card = document.createElement('div');
      card.className = 'card' + (it.id === selectedId ? ' selected' : '');
      card.draggable = true;
      card.dataset.id = String(it.id);
      card.innerHTML =
        '<div class="card-head"><span class="card-num">#' + (pos + 1) + '</span>' +
        '<span class="card-tools"><button data-a="rot" title="Повернуть на 90° по часовой">↻</button>' +
        '<button data-a="del" title="Удалить из списка">×</button></span></div>' +
        '<div class="thumb"></div>' +
        '<div class="card-foot"><span class="card-name"></span><span class="card-badge"></span></div>';
      card.querySelector('.card-name').textContent = it.name;
      card.querySelector('.card-name').title = it.name;
      card.querySelector('.card-badge').textContent = isChanged(it) ? cardSizeText(it) : '';
      const cv = document.createElement('canvas');
      drawThumb(it, cv);
      card.querySelector('.thumb').appendChild(cv);

      card.addEventListener('click', (e) => {
        if (e.target.closest('button')) return;
        select(it.id);
      });
      card.querySelector('[data-a="rot"]').addEventListener('click', (e) => { e.stopPropagation(); rotate(it, 90); });
      card.querySelector('[data-a="del"]').addEventListener('click', (e) => { e.stopPropagation(); removeItem(it, card); });
      setupCardDrag(card, it.id);
      cardsEl.appendChild(card);
    });
  }

  function refreshCard(it) {
    const card = cardsEl.querySelector('.card[data-id="' + it.id + '"]');
    if (!card) return;
    drawThumb(it, card.querySelector('canvas'));
    card.querySelector('.card-badge').textContent = isChanged(it) ? cardSizeText(it) : '';
  }

  const current = () => items.find((i) => i.id === selectedId) || null;

  function select(id) {
    if (selectedId === id) return;
    selectedId = id;
    cardsEl.querySelectorAll('.card').forEach((c) => c.classList.toggle('selected', c.dataset.id === String(id)));
    renderStage();
  }

  function renderStage() {
    const it = current();
    $('empty').hidden = !!it;
    cropWrap.hidden = !it;
    $('stageBar').hidden = !it;
    if (!it) return;
    const rd = rotatedDims(it);
    const host = $('stageInner');
    const availW = Math.max(120, host.clientWidth - 72);
    const availH = Math.max(120, host.clientHeight - 72);
    const k = Math.min(availW / rd.w, availH / rd.h, 1.6);
    const cssW = Math.max(40, Math.round(rd.w * k)), cssH = Math.max(40, Math.round(rd.h * k));
    const dpr = window.devicePixelRatio || 1;
    stageCanvas.style.width = cssW + 'px'; stageCanvas.style.height = cssH + 'px';
    stageCanvas.width = Math.round(cssW * dpr); stageCanvas.height = Math.round(cssH * dpr);
    const ctx = stageCanvas.getContext('2d');
    ctx.clearRect(0, 0, stageCanvas.width, stageCanvas.height);
    paint(ctx, it.pv, it.pw, it.ph, it.rot, FULL(), stageCanvas.width, stageCanvas.height);
    positionCrop(it);
  }

  function positionCrop(it) {
    const c = it.crop;
    cropEl.style.left = (c.x * 100) + '%';
    cropEl.style.top = (c.y * 100) + '%';
    cropEl.style.width = (c.w * 100) + '%';
    cropEl.style.height = (c.h * 100) + '%';
    const rd = rotatedDims(it);
    $('cropInfo').textContent = Math.round(c.w * rd.w) + ' × ' + Math.round(c.h * rd.h) + ' px';
  }

  // ---------- действия над фото ----------
  function rotate(it, deg) {
    const c = it.crop;
    if (deg === 90) it.crop = { x: 1 - (c.y + c.h), y: c.x, w: c.h, h: c.w };
    else it.crop = { x: c.y, y: 1 - (c.x + c.w), w: c.h, h: c.w };
    it.rot = (it.rot + (deg === 90 ? 90 : 270)) % 360;
    refreshCard(it);
    if (it.id === selectedId) renderStage();
    updateInfo();
  }

  function removeItem(it, card) {
    if (!api.confirm('Убрать «' + it.name + '» из списка?')) return;
    const idx = items.findIndex((i) => i.id === it.id);
    const finish = () => {
      items.splice(idx, 1);
      if (selectedId === it.id) selectedId = items.length ? items[Math.min(idx, items.length - 1)].id : null;
      renderAll();
    };
    if (card) { card.classList.add('removing'); setTimeout(finish, 150); } else finish();
  }

  $('btnRotL').addEventListener('click', () => { const it = current(); if (it) rotate(it, -90); });
  $('btnRotR').addEventListener('click', () => { const it = current(); if (it) rotate(it, 90); });
  $('btnCropReset').addEventListener('click', () => {
    const it = current(); if (!it) return;
    it.crop = FULL(); positionCrop(it); refreshCard(it); updateInfo();
  });

  // ---------- рамка обрезки ----------
  let drag = null;
  cropEl.addEventListener('pointerdown', (e) => {
    const it = current(); if (!it) return;
    const h = e.target.closest('[data-h]');
    drag = { mode: h ? h.dataset.h : 'move', sx: e.clientX, sy: e.clientY, start: { ...it.crop }, id: it.id };
    cropEl.setPointerCapture(e.pointerId);
    e.preventDefault();
  });
  cropEl.addEventListener('pointermove', (e) => {
    if (!drag) return;
    const it = current(); if (!it || it.id !== drag.id) return;
    const r = cropWrap.getBoundingClientRect();
    const dx = (e.clientX - drag.sx) / r.width, dy = (e.clientY - drag.sy) / r.height;
    let { x, y, w, h } = drag.start;
    let x2 = x + w, y2 = y + h;
    if (drag.mode === 'move') {
      x = clamp(x + dx, 0, 1 - w); y = clamp(y + dy, 0, 1 - h); x2 = x + w; y2 = y + h;
    } else {
      if (drag.mode.includes('w')) x = clamp(x + dx, 0, x2 - MIN_CROP);
      if (drag.mode.includes('e')) x2 = clamp(x2 + dx, x + MIN_CROP, 1);
      if (drag.mode.includes('n')) y = clamp(y + dy, 0, y2 - MIN_CROP);
      if (drag.mode.includes('s')) y2 = clamp(y2 + dy, y + MIN_CROP, 1);
    }
    it.crop = { x, y, w: x2 - x, h: y2 - y };
    positionCrop(it);
  });
  const endDrag = (e) => {
    if (!drag) return;
    const it = current();
    drag = null;
    try { cropEl.releasePointerCapture(e.pointerId); } catch {}
    if (it) { refreshCard(it); updateInfo(); }
  };
  cropEl.addEventListener('pointerup', endDrag);
  cropEl.addEventListener('pointercancel', endDrag);

  window.addEventListener('resize', () => { if (current()) renderStage(); });

  // ---------- порядок карточек ----------
  let dragSrcId = null;
  function setupCardDrag(card, id) {
    card.addEventListener('dragstart', (e) => {
      dragSrcId = id;
      card.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/x-qd-card', String(id));
    });
    card.addEventListener('dragend', () => {
      dragSrcId = null;
      card.classList.remove('dragging');
      cardsEl.querySelectorAll('.card').forEach((c) => c.classList.remove('drag-over'));
    });
    card.addEventListener('dragover', (e) => {
      if (dragSrcId === null) return; // это перетаскивание файла — обработает окно
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      if (dragSrcId !== id) card.classList.add('drag-over');
    });
    card.addEventListener('dragleave', () => card.classList.remove('drag-over'));
    card.addEventListener('drop', (e) => {
      if (dragSrcId === null) return;
      e.preventDefault(); e.stopPropagation();
      card.classList.remove('drag-over');
      if (dragSrcId === id) return;
      const from = items.findIndex((i) => i.id === dragSrcId);
      const to = items.findIndex((i) => i.id === id);
      const [moved] = items.splice(from, 1);
      items.splice(to, 0, moved);
      dragSrcId = null;
      renderCards();
    });
  }

  // ---------- добавление: перетаскивание файлов в окно ----------
  const overlay = $('dropOverlay');
  let dz = 0;
  const hasFiles = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
  window.addEventListener('dragenter', (e) => { if (!hasFiles(e)) return; e.preventDefault(); dz++; overlay.classList.add('active'); });
  window.addEventListener('dragover', (e) => { if (!hasFiles(e)) return; e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; });
  window.addEventListener('dragleave', (e) => { if (!hasFiles(e)) return; dz--; if (dz <= 0) { dz = 0; overlay.classList.remove('active'); } });
  window.addEventListener('drop', async (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); dz = 0; overlay.classList.remove('active');
    await addFiles(e.dataTransfer.files);
  });

  // ---------- добавление: кнопка «Обзор» ----------
  async function browse() {
    const r = await api.pick();
    if (!r.ok) { setStatus('Ошибка: ' + r.error, 'error'); return; }
    await addPaths(r.paths);
  }
  $('btnAdd').addEventListener('click', browse);
  $('btnAddCard').addEventListener('click', browse);
  $('btnAddEmpty').addEventListener('click', browse);

  // ---------- сброс ----------
  $('btnReset').addEventListener('click', () => {
    if (!api.confirm('Сбросить обрезку и повороты у всех фото?')) return;
    items.forEach((it) => { it.rot = 0; it.crop = FULL(); });
    renderAll();
    setStatus('Изменения сброшены', 'warning');
  });

  // ---------- экспорт в PDF ----------
  async function renderItem(it) {
    // Нетронутый JPEG с обычной ориентацией вставляем как есть — без повторного сжатия.
    if (it.mime === 'image/jpeg' && it.exif === 1 && it.rot === 0 && isFull(it.crop)) {
      return { bytes: it.bytes, mime: 'image/jpeg' };
    }
    const bmp = await createImageBitmap(new Blob([it.bytes], { type: it.mime }), { imageOrientation: 'from-image' });
    const swap = it.rot % 180 !== 0;
    const rw = swap ? bmp.height : bmp.width, rh = swap ? bmp.width : bmp.height;
    const cw = Math.max(1, Math.round(it.crop.w * rw)), ch = Math.max(1, Math.round(it.crop.h * rh));
    const canvas = document.createElement('canvas');
    canvas.width = cw; canvas.height = ch;
    const ctx = canvas.getContext('2d');
    const outMime = (it.mime === 'image/jpeg' || it.mime === 'image/bmp') ? 'image/jpeg' : 'image/png';
    if (outMime === 'image/jpeg') { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cw, ch); }
    paint(ctx, bmp, bmp.width, bmp.height, it.rot, it.crop, cw, ch);
    bmp.close();
    const blob = await new Promise((res) => canvas.toBlob(res, outMime, 0.93));
    canvas.width = canvas.height = 0;
    if (!blob) throw new Error('Не удалось подготовить «' + it.name + '»');
    return { bytes: new Uint8Array(await blob.arrayBuffer()), mime: outMime };
  }

  $('btnSave').addEventListener('click', async () => {
    if (busy || !items.length) return;
    busy = true; updateInfo();
    try {
      await api.stageReset();
      for (let i = 0; i < items.length; i++) {
        setStatus('Подготовка ' + (i + 1) + ' из ' + items.length + '…', '');
        const page = await renderItem(items[i]);
        await api.stagePage(page);
      }
      setStatus('Сборка PDF…', '');
      const r = await api.finish({ pageMode: $('pageMode').value, suggestedName: baseName(items[0].name) });
      if (r.ok) setStatus('Сохранено: ' + r.path, 'success');
      else if (r.canceled) setStatus('Сохранение отменено', 'warning');
      else setStatus('Ошибка: ' + r.error, 'error');
    } catch (e) {
      setStatus('Ошибка: ' + (e && e.message || e), 'error');
    } finally {
      busy = false; updateInfo();
    }
  });

  api.onAdd((p) => { if (p && p.paths) addPaths(p.paths); });
  renderAll();
})();
