(function () {
  const img = document.getElementById('iconImg');
  const { qd } = window;
  const ICON_SOURCES = { normal: '../../build/icon-normal.png', blink: '../../build/icon-normal-blink.png', file: '../../build/icon-file.png' };
  let currentState = 'normal';
  function applyState(state) {
    if (state === currentState) return;
    currentState = state;
    img.classList.remove('state-normal', 'state-blink', 'state-file');
    img.classList.add('state-' + state);
    const src = ICON_SOURCES[state] || ICON_SOURCES.normal;
    const fullSrc = new URL(src, document.baseURI).href;
    if (img.getAttribute('src') !== fullSrc) img.src = fullSrc;
  }
  applyState('normal');
  qd.icon.onSetState((state) => applyState(state));
  let isHovering = false;
  function setHover(val) { if (isHovering === val) return; isHovering = val; qd.icon.setHover(val); }
  document.addEventListener('mouseenter', () => setHover(true));
  document.addEventListener('mouseover', () => setHover(true));
  document.addEventListener('mousemove', () => setHover(true));
  document.addEventListener('mouseleave', () => setHover(false));
  let dragging = false, lastScreenX = 0, lastScreenY = 0;
  document.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    if (document.body.classList.contains('drop-active')) return;
    dragging = true; lastScreenX = e.screenX; lastScreenY = e.screenY;
    document.body.classList.add('dragging'); qd.icon.dragStart(); e.preventDefault();
  });
  document.addEventListener('mousemove', (e) => {
    if (!dragging) return;
    const dx = e.screenX - lastScreenX, dy = e.screenY - lastScreenY;
    if (dx !== 0 || dy !== 0) { qd.icon.dragMove(dx, dy); lastScreenX = e.screenX; lastScreenY = e.screenY; }
  });
  function endDrag() { if (!dragging) return; dragging = false; document.body.classList.remove('dragging'); qd.icon.dragEnd(); }
  document.addEventListener('mouseup', endDrag);
  window.addEventListener('blur', endDrag);
  let dragCounter = 0;
  document.addEventListener('dragenter', (e) => {
    if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return;
    e.preventDefault(); dragCounter++; document.body.classList.add('drop-active'); qd.icon.setFileHover(true);
  });
  document.addEventListener('dragover', (e) => {
    if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return;
    e.preventDefault(); e.dataTransfer.dropEffect = 'copy';
  });
  document.addEventListener('dragleave', () => { dragCounter--; if (dragCounter <= 0) { dragCounter = 0; document.body.classList.remove('drop-active'); qd.icon.setFileHover(false); } });
  document.addEventListener('drop', (e) => {
    e.preventDefault(); dragCounter = 0; document.body.classList.remove('drop-active'); qd.icon.setFileHover(false);
    const files = Array.from(e.dataTransfer.files || []);
    if (files.length === 0) return;
    const paths = files.map((f) => qd.files.pathForFile(f)).filter(Boolean);
    if (paths.length > 0) qd.files.add(paths).then(() => qd.icon.setHover(true));
  });
  window.addEventListener('contextmenu', (e) => e.preventDefault());
})();
