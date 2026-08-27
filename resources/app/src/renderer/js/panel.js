(function () {
  const { qd } = window;
  document.body.classList.add('collapsed');
  qd.panel.onSetExpanded((expanded) => {
    if (expanded) { document.body.classList.remove('collapsed'); document.body.classList.add('expanded'); }
    else { document.body.classList.remove('expanded'); document.body.classList.add('collapsed'); }
  });
  document.addEventListener('mousedown', (e) => {
    if (e.target.tagName === 'TEXTAREA' || e.target.tagName === 'INPUT') qd.panel.focus();
  }, true);
  let panelHoverActive = false;
  document.addEventListener('mouseenter', () => { if (!panelHoverActive) { panelHoverActive = true; qd.panel.setHover(true); } });
  document.addEventListener('mouseleave', () => { if (panelHoverActive) { panelHoverActive = false; setTimeout(() => { if (!panelHoverActive) qd.panel.setHover(false); }, 50); } });
  const tabs = document.querySelectorAll('.tab');
  const bodies = document.querySelectorAll('.tab-body');
  tabs.forEach((tab) => tab.addEventListener('click', () => switchTab(tab.dataset.tab)));
  function switchTab(name) { tabs.forEach((t) => t.classList.toggle('active', t.dataset.tab === name)); bodies.forEach((b) => b.classList.toggle('active', b.dataset.tab === name)); }
  let clipboardHistory = [];
  function renderClipboard(history) {
    clipboardHistory = history || [];
    const list = document.getElementById('clipList');
    if (clipboardHistory.length === 0) { list.innerHTML = '<li class="empty">История пуста — скопируйте что-нибудь.</li>'; return; }
    list.innerHTML = '';
    clipboardHistory.forEach((text, i) => {
      const li = document.createElement('li');
      li.className = 'clip-item' + (i === 0 ? ' current' : '');
      const num = String(i + 1).padStart(2, '0');
      const preview = text.length <= 240 ? text : text.slice(0, 110) + ' … ' + text.slice(-110);
      li.innerHTML = '<span class="clip-meta">' + num + '</span><div class="clip-text">' + escapeHtml(preview) + '</div><span class="clip-action"></span>';
      li.addEventListener('click', () => { qd.clipboard.copyBack(text); flashFooter('Скопировано в буфер'); });
      list.appendChild(li);
    });
  }
  qd.clipboard.onUpdated(renderClipboard);
  qd.clipboard.get().then(renderClipboard);
  let filesList = [];
  const selectedPaths = new Set();
  function renderFiles() {
    const list = document.getElementById('fileList');
    document.getElementById('filesCount').textContent = filesList.length + ' файл' + (filesList.length === 1 ? '' : 'ов');
    if (filesList.length === 0) { list.innerHTML = '<li class="empty">Файлов нет.</li>'; updateBuildBtn(); return; }
    list.innerHTML = '';
    filesList.forEach((file) => {
      const li = document.createElement('li');
      li.className = 'file-item' + (selectedPaths.has(file.path) ? ' selected' : '') + (file.isOutput ? ' is-output' : '');
      li.dataset.path = file.path;
      li.innerHTML = '<span class="file-grip" title="Перетащите, чтобы вытащить файл" draggable="true">⠿</span><span class="file-check"></span><span class="file-icon">' + iconForExt(file.ext) + '</span><div class="file-info"><span class="file-name">' + escapeHtml(file.name) + '</span><span class="file-meta">' + file.ext.replace('.', '').toUpperCase() + ' · ' + formatSize(file.size) + '</span></div><div class="file-actions"><button class="ghost small" data-action="open" title="Показать в проводнике">⤢</button><button class="ghost small" data-action="remove" title="Удалить">✕</button></div>';
      li.addEventListener('click', (e) => { if (e.target.closest('button') || e.target.closest('.file-grip')) return; toggleSelect(file.path); });
      li.querySelector('[data-action="open"]').addEventListener('click', (e) => { e.stopPropagation(); qd.files.openInExplorer(file.path); });
      li.querySelector('[data-action="remove"]').addEventListener('click', async (e) => { e.stopPropagation(); await qd.files.remove(file.path); selectedPaths.delete(file.path); filesList = await qd.files.list(); renderFiles(); });
      setupDragOut(li.querySelector('.file-grip'), file);
      list.appendChild(li);
    });
    updateBuildBtn();
  }
  function toggleSelect(path) { if (selectedPaths.has(path)) selectedPaths.delete(path); else selectedPaths.add(path); renderFiles(); }
  function updateBuildBtn() { document.getElementById('btnBuildPdf').disabled = selectedPaths.size === 0; }
  function setupDragOut(gripEl, file) {
    const emptyImg = new Image();
    emptyImg.src = 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
    gripEl.addEventListener('dragstart', (e) => { e.dataTransfer.setDragImage(emptyImg, 0, 0); e.dataTransfer.effectAllowed = 'copyMove'; e.dataTransfer.setData('text/plain', file.path); qd.files.dragOut(file.path, null); });
  }
  const dz = document.getElementById('dropzone');
  const filePicker = document.getElementById('filePicker');
  let dzCounter = 0;
  dz.addEventListener('click', (e) => { if (e.target.closest('.dz-pick')) return; filePicker.click(); });
  dz.addEventListener('dragenter', (e) => { if (!hasFiles(e)) return; e.preventDefault(); dzCounter++; dz.classList.add('active'); });
  dz.addEventListener('dragover', (e) => { if (!hasFiles(e)) return; e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; });
  dz.addEventListener('dragleave', () => { dzCounter--; if (dzCounter <= 0) { dzCounter = 0; dz.classList.remove('active'); } });
  dz.addEventListener('drop', async (e) => {
    e.preventDefault(); dzCounter = 0; dz.classList.remove('active');
    const files = Array.from(e.dataTransfer.files || []);
    const paths = files.map((f) => qd.files.pathForFile(f)).filter(Boolean);
    if (paths.length > 0) { await qd.files.add(paths); flashFooter('Добавлено: ' + paths.length); }
    else flashFooter('Не удалось определить путь файла');
  });
  filePicker.addEventListener('change', async () => {
    const files = Array.from(filePicker.files);
    const paths = files.map((f) => qd.files.pathForFile(f)).filter(Boolean);
    if (paths.length > 0) await qd.files.add(paths);
    filePicker.value = '';
  });
  function hasFiles(e) { return Array.from(e.dataTransfer.types || []).includes('Files'); }
  document.getElementById('btnBuildPdf').addEventListener('click', async () => {
    const paths = Array.from(selectedPaths);
    if (paths.length === 0) return;
    const status = document.getElementById('pdfStatus');
    status.className = 'pdf-status'; status.textContent = 'Сборка PDF…';
    try {
      const result = await qd.files.buildPdf(paths);
      if (result.errors && result.errors.length > 0) { status.className = 'pdf-status error'; status.textContent = 'Готово с ошибками: ' + result.errors.slice(0, 2).join('; '); }
      else { status.className = 'pdf-status ok'; status.textContent = '✓ ' + result.outPath.split(/[\\/]/).pop(); }
      filesList = result.files; selectedPaths.clear(); selectedPaths.add(result.outPath); renderFiles();
    } catch (e) { status.className = 'pdf-status error'; status.textContent = 'Ошибка: ' + e.message; }
  });
  document.getElementById('btnSelectAll').addEventListener('click', () => {
    const inputs = filesList.filter((f) => !f.isOutput);
    if (inputs.length === 0) return;
    const allSelected = inputs.every((f) => selectedPaths.has(f.path));
    if (allSelected) inputs.forEach((f) => selectedPaths.delete(f.path));
    else inputs.forEach((f) => selectedPaths.add(f.path));
    renderFiles();
  });
  document.getElementById('btnClearFiles').addEventListener('click', async () => { await qd.files.clear(); filesList = []; selectedPaths.clear(); renderFiles(); document.getElementById('pdfStatus').textContent = ''; });
  qd.files.onAdded((files) => { filesList = files; renderFiles(); });
  qd.files.list().then((files) => { filesList = files; renderFiles(); });
  let replies = [];
  let editingIdx = -1;
  function renderReplies() {
    const list = document.getElementById('repliesList');
    if (replies.length === 0) { list.innerHTML = '<li class="empty">Список пуст.</li>'; return; }
    list.innerHTML = '';
    replies.forEach((text, i) => {
      const li = document.createElement('li'); li.className = 'reply-item';
      const num = String(i + 1).padStart(2, '0');
      if (editingIdx === i) {
        const ta = document.createElement('textarea'); ta.value = text;
        const saveBtn = document.createElement('button'); saveBtn.className = 'primary small'; saveBtn.textContent = 'сохранить';
        saveBtn.addEventListener('click', async () => { replies[i] = ta.value; editingIdx = -1; await qd.replies.save(replies); renderReplies(); });
        const cancelBtn = document.createElement('button'); cancelBtn.className = 'ghost small'; cancelBtn.textContent = 'отмена';
        cancelBtn.addEventListener('click', () => { editingIdx = -1; renderReplies(); });
        const actions = document.createElement('div'); actions.className = 'reply-edit-actions'; actions.appendChild(cancelBtn); actions.appendChild(saveBtn);
        const wrap = document.createElement('div'); wrap.className = 'reply-edit'; wrap.appendChild(ta); wrap.appendChild(actions);
        li.appendChild(wrap);
        setTimeout(() => { ta.focus(); ta.selectionStart = ta.value.length; }, 0);
      } else {
        const numEl = document.createElement('span'); numEl.className = 'reply-num'; numEl.textContent = num;
        const txt = document.createElement('div'); txt.className = 'reply-text'; txt.textContent = text.length <= 240 ? text : text.slice(0, 110) + ' … ' + text.slice(-110);
        txt.addEventListener('click', async () => { await qd.clipboard.copyBack(text); flashFooter('Скопировано в буфер'); });
        const actions = document.createElement('div'); actions.className = 'reply-actions';
        const editBtn = document.createElement('button'); editBtn.className = 'icon-btn ghost'; editBtn.textContent = '✎'; editBtn.title = 'Изменить';
        editBtn.addEventListener('click', () => { editingIdx = i; renderReplies(); });
        const rmBtn = document.createElement('button'); rmBtn.className = 'icon-btn ghost'; rmBtn.textContent = '✕'; rmBtn.title = 'Удалить';
        rmBtn.addEventListener('click', async () => { replies.splice(i, 1); await qd.replies.save(replies); renderReplies(); });
        actions.appendChild(editBtn); actions.appendChild(rmBtn);
        li.appendChild(numEl); li.appendChild(txt); li.appendChild(actions);
      }
      list.appendChild(li);
    });
  }
  document.getElementById('btnAddReply').addEventListener('click', async () => {
    const input = document.getElementById('replyInput');
    const text = input.value.trim();
    if (!text) return;
    replies.push(text); await qd.replies.save(replies); input.value = ''; renderReplies(); flashFooter('Добавлено');
  });
  document.getElementById('replyInput').addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); document.getElementById('btnAddReply').click(); }
  });
  qd.replies.get().then((list) => { replies = list || []; renderReplies(); });
  document.addEventListener('dragover', (e) => { if (!e.target.closest('#dropzone, .file-grip')) e.preventDefault(); });
  document.addEventListener('drop', (e) => { if (!e.target.closest('#dropzone, .file-grip')) e.preventDefault(); });
  function escapeHtml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function formatSize(b) { if (b < 1024) return b + ' B'; if (b < 1048576) return (b / 1024).toFixed(1) + ' KB'; return (b / 1048576).toFixed(1) + ' MB'; }
  function iconForExt(ext) {
    const e = ext.toLowerCase();
    if (['.png','.jpg','.jpeg','.jfif','.bmp','.gif','.webp'].includes(e)) return 'IMG';
    if (e === '.docx' || e === '.doc') return 'DOC';
    if (e === '.xlsx' || e === '.xls') return 'XLS';
    if (e === '.pdf') return 'PDF';
    return '?';
  }
  function flashFooter(msg) {
    const footer = document.getElementById('footerInfo');
    const prev = footer.textContent;
    footer.textContent = msg;
    setTimeout(() => { footer.textContent = prev; }, 1400);
  }
})();
