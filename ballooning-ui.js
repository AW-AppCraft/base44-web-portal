/* ballooning-ui.js — viewer, balloon overlay and characteristic table.
 * Engine lives in ballooning.js; this file only handles interaction and export
 * plumbing. State persists to localStorage so a half-ballooned drawing survives
 * a reload. */

(function () {
  'use strict';

  var B = window.Ballooning;
  var STORAGE_KEY = 'gdt-ballooning-project';

  var state = {
    loader: null,
    pages: [],
    pageIndex: 0,
    scale: 1.4,
    chars: [],
    selectedId: null,
    fileName: '',
    addMode: false,
    meta: {},
    opts: {
      unit: 'mm',
      generalTol: { one: 0.5, two: 0.25, three: 0.1, angular: 0.5 },
      zoneRows: 0,
      zoneCols: 0,
      includeNotes: true,
      symbolMap: B.SYMBOL_FONT_MAP
    }
  };

  var el = {};

  function $(id) { return document.getElementById(id); }

  function init() {
    ['fileInput', 'viewer', 'pdfCanvas', 'overlay', 'charTable', 'charBody', 'summary',
     'pageLabel', 'prevPage', 'nextPage', 'zoomIn', 'zoomOut', 'status', 'reprocess',
     'addBalloon', 'renumber', 'clearAll', 'unit', 'tolOne', 'tolTwo', 'tolThree',
     'tolAngular', 'zoneRows', 'zoneCols', 'includeNotes', 'partNumber', 'partName',
     'drawingNumber', 'revision', 'fairNumber', 'serialNumber', 'supplier', 'sampleSize',
     'exportForm3', 'exportPpap', 'exportJson', 'exportPng', 'saveLocal', 'loadLocal',
     'detail', 'emptyState'].forEach(function (k) { el[k] = $(k); });

    el.fileInput.addEventListener('change', onFile);
    el.prevPage.addEventListener('click', function () { gotoPage(state.pageIndex - 1); });
    el.nextPage.addEventListener('click', function () { gotoPage(state.pageIndex + 1); });
    el.zoomIn.addEventListener('click', function () { setScale(state.scale * 1.25); });
    el.zoomOut.addEventListener('click', function () { setScale(state.scale / 1.25); });
    el.reprocess.addEventListener('click', function () { runExtraction(true); });
    el.addBalloon.addEventListener('click', toggleAddMode);
    el.renumber.addEventListener('click', function () { B.assignNumbers(state.chars); renderAll(); });
    el.clearAll.addEventListener('click', clearAll);
    el.overlay.addEventListener('click', onOverlayClick);

    el.exportForm3.addEventListener('click', function () {
      download(fileBase() + '-AS9102-Form3.csv', B.exportAS9102Form3(sorted(), readMeta()), 'text/csv');
    });
    el.exportPpap.addEventListener('click', function () {
      download(fileBase() + '-PPAP-dimensional-results.csv',
        B.exportPPAP(sorted(), readMeta(), parseInt(el.sampleSize.value, 10) || 5), 'text/csv');
    });
    el.exportJson.addEventListener('click', function () {
      download(fileBase() + '-ballooning.json', B.exportJson(sorted(), readMeta()), 'application/json');
    });
    el.exportPng.addEventListener('click', exportBalloonedPng);
    el.saveLocal.addEventListener('click', saveLocal);
    el.loadLocal.addEventListener('click', loadLocal);

    window.addEventListener('resize', positionOverlay);
    setStatus('Load a drawing PDF to begin.');
    if (localStorage.getItem(STORAGE_KEY)) {
      setStatus('Load a drawing PDF to begin. A saved characteristic list is available — load the same PDF, then click "Load saved".');
    }
  }

  // -------------------------------------------------------------------------
  // Loading and extraction
  // -------------------------------------------------------------------------

  function onFile(e) {
    var file = e.target.files && e.target.files[0];
    if (!file) return;
    if (!window.pdfjsLib) {
      setStatus('PDF.js has not finished loading yet — wait a moment and pick the file again.', true);
      return;
    }
    state.fileName = file.name;
    setStatus('Reading ' + file.name + ' …');

    var reader = new FileReader();
    reader.onload = function () {
      state.loader = new B.PdfLoader(window.pdfjsLib);
      state.loader.load(reader.result).then(function (pages) {
        state.pages = pages;
        state.pageIndex = 0;
        el.emptyState.style.display = 'none';
        runExtraction(false);
      }).catch(function (err) {
        setStatus('Could not read PDF: ' + err.message, true);
      });
    };
    reader.readAsArrayBuffer(file);
  }

  function readOpts() {
    state.opts.unit = el.unit.value;
    state.opts.generalTol = {
      one: parseFloat(el.tolOne.value) || 0,
      two: parseFloat(el.tolTwo.value) || 0,
      three: parseFloat(el.tolThree.value) || 0,
      angular: parseFloat(el.tolAngular.value) || 0
    };
    state.opts.zoneRows = parseInt(el.zoneRows.value, 10) || 0;
    state.opts.zoneCols = parseInt(el.zoneCols.value, 10) || 0;
    state.opts.includeNotes = el.includeNotes.checked;
    return state.opts;
  }

  function runExtraction(keepManual) {
    if (!state.pages.length) { setStatus('Load a PDF first.', true); return; }
    var manual = keepManual ? state.chars.filter(function (c) { return c.source === 'manual'; }) : [];
    var textCount = state.pages.reduce(function (n, p) { return n + p.items.length; }, 0);

    if (textCount === 0) {
      state.chars = manual;
      setStatus('No text layer found — this PDF is a scan or has outlined text. ' +
        'Automatic extraction is not possible; balloon manually with "+ Add balloon", ' +
        'or run the PDF through OCR / a vision model first.', true);
    } else {
      state.chars = B.extract(state.pages, readOpts()).concat(manual);
      B.assignNumbers(state.chars);
      var s = B.summarize(state.chars);
      setStatus('Extracted ' + s.total + ' characteristics from ' + textCount + ' text objects. ' +
        s.review + ' need review. Verify every one against the print before submitting a FAIR.');
    }
    renderPage();
  }

  // -------------------------------------------------------------------------
  // Rendering
  // -------------------------------------------------------------------------

  function renderPage() {
    if (!state.pages.length) return;
    state.pageIndex = Math.max(0, Math.min(state.pageIndex, state.pages.length - 1));
    el.pageLabel.textContent = 'Sheet ' + (state.pageIndex + 1) + ' / ' + state.pages.length;
    state.loader.renderPage(state.pageIndex, el.pdfCanvas, state.scale).then(renderAll);
  }

  function renderAll() {
    positionOverlay();
    renderOverlay();
    renderTable();
    renderSummary();
  }

  function positionOverlay() {
    el.overlay.style.width = el.pdfCanvas.width + 'px';
    el.overlay.style.height = el.pdfCanvas.height + 'px';
  }

  function pageChars() {
    return state.chars.filter(function (c) { return c.pageIndex === state.pageIndex; });
  }

  function renderOverlay() {
    var W = el.pdfCanvas.width, H = el.pdfCanvas.height;
    el.overlay.innerHTML = '';

    var svgNS = 'http://www.w3.org/2000/svg';
    var svg = document.createElementNS(svgNS, 'svg');
    svg.setAttribute('class', 'leaders');
    svg.setAttribute('width', W);
    svg.setAttribute('height', H);
    el.overlay.appendChild(svg);

    pageChars().forEach(function (c) {
      var bx = c.balloon.x * W, by = c.balloon.y * H;

      if (c.box) {
        var hi = document.createElement('div');
        hi.className = 'hl hl-' + typeClass(c.type) + (c.id === state.selectedId ? ' sel' : '');
        hi.style.left = (c.box.x * W - 2) + 'px';
        hi.style.top = (c.box.y * H - 2) + 'px';
        hi.style.width = (c.box.w * W + 4) + 'px';
        hi.style.height = (c.box.h * H + 4) + 'px';
        hi.dataset.id = c.id;
        el.overlay.appendChild(hi);

        var line = document.createElementNS(svgNS, 'line');
        line.setAttribute('x1', bx); line.setAttribute('y1', by);
        line.setAttribute('x2', (c.box.x + c.box.w / 2) * W);
        line.setAttribute('y2', (c.box.y + c.box.h / 2) * H);
        line.setAttribute('class', 'leader' + (c.id === state.selectedId ? ' sel' : ''));
        svg.appendChild(line);
      }

      var b = document.createElement('div');
      b.className = 'balloon b-' + typeClass(c.type) +
        (c.id === state.selectedId ? ' sel' : '') + (c.review ? ' review' : '');
      b.style.left = bx + 'px';
      b.style.top = by + 'px';
      b.textContent = c.number;
      b.title = c.type + ' — ' + c.text;
      b.dataset.id = c.id;
      makeDraggable(b, c);
      el.overlay.appendChild(b);
    });
  }

  function typeClass(t) {
    return ({ 'Dimension': 'dim', 'GD&T': 'gdt', 'Material': 'mat', 'Coating': 'coat',
              'Thread': 'thr', 'Surface Finish': 'sf', 'Note': 'note', 'Process': 'proc' })[t] || 'other';
  }

  function makeDraggable(node, c) {
    var dragging = false, moved = false, sx = 0, sy = 0, ox = 0, oy = 0;

    node.addEventListener('mousedown', function (e) {
      dragging = true; moved = false;
      sx = e.clientX; sy = e.clientY; ox = c.balloon.x; oy = c.balloon.y;
      e.preventDefault(); e.stopPropagation();
    });

    document.addEventListener('mousemove', function (e) {
      if (!dragging) return;
      var dx = (e.clientX - sx) / el.pdfCanvas.width;
      var dy = (e.clientY - sy) / el.pdfCanvas.height;
      if (Math.abs(e.clientX - sx) + Math.abs(e.clientY - sy) > 3) moved = true;
      c.balloon.x = Math.max(0, Math.min(1, ox + dx));
      c.balloon.y = Math.max(0, Math.min(1, oy + dy));
      renderOverlay();
    });

    document.addEventListener('mouseup', function () {
      if (!dragging) return;
      dragging = false;
      if (!moved) select(c.id);
    });
  }

  function onOverlayClick(e) {
    if (state.addMode) {
      var rect = el.pdfCanvas.getBoundingClientRect();
      addManual((e.clientX - rect.left) / rect.width, (e.clientY - rect.top) / rect.height);
      return;
    }
    var id = e.target && e.target.dataset && e.target.dataset.id;
    if (id) select(id);
  }

  function toggleAddMode() {
    if (!state.pages.length) { setStatus('Load a PDF first.', true); return; }
    state.addMode = !state.addMode;
    el.addBalloon.classList.toggle('active', state.addMode);
    el.overlay.classList.toggle('adding', state.addMode);
    setStatus(state.addMode
      ? 'Click anywhere on the drawing to place a manual balloon (for callouts the parser missed).'
      : 'Manual placement off.');
  }

  function addManual(nx, ny) {
    var page = state.pages[state.pageIndex];
    var c = {
      id: B.uid(),
      pageIndex: state.pageIndex,
      pageNumber: page.pageNumber,
      nx: nx, ny: ny,
      box: null,
      balloon: { x: nx, y: ny },
      type: 'Dimension',
      subtype: 'Manually added',
      text: '',
      requirement: '',
      nominal: null, upper: null, lower: null,
      unit: state.opts.unit,
      method: '',
      zone: B.computeZone(nx, ny, state.opts.zoneRows, state.opts.zoneCols),
      flags: ['Manually added'],
      review: true,
      critical: false,
      source: 'manual'
    };
    state.chars.push(c);
    B.assignNumbers(state.chars);
    state.selectedId = c.id;
    renderAll();
    setStatus('Manual balloon added — fill in the requirement in the table below.');
  }

  function select(id) {
    state.selectedId = id;
    renderOverlay();
    renderTable();
    var row = document.querySelector('tr[data-id="' + id + '"]');
    if (row) row.scrollIntoView({ block: 'nearest' });
  }

  // -------------------------------------------------------------------------
  // Characteristic table
  // -------------------------------------------------------------------------

  function sorted() {
    return state.chars.slice().sort(function (a, b) { return a.number - b.number; });
  }

  function renderTable() {
    var rows = sorted();
    el.charBody.innerHTML = '';

    rows.forEach(function (c) {
      var tr = document.createElement('tr');
      tr.dataset.id = c.id;
      if (c.id === state.selectedId) tr.className = 'sel';
      if (c.review) tr.classList.add('needs-review');

      tr.appendChild(cell(String(c.number), 'num'));
      tr.appendChild(cell('Sh' + c.pageNumber + (c.zone ? ' / ' + c.zone : ''), 'loc'));
      tr.appendChild(selectCell(c, 'type',
        ['Dimension', 'GD&T', 'Thread', 'Surface Finish', 'Material', 'Coating', 'Process', 'Note']));
      tr.appendChild(editCell(c, 'requirement', 'req'));
      tr.appendChild(numCell(c, 'nominal'));
      tr.appendChild(numCell(c, 'lower'));
      tr.appendChild(numCell(c, 'upper'));
      tr.appendChild(editCell(c, 'method', 'method'));
      tr.appendChild(checkCell(c, 'critical'));

      var flags = cell(c.flags.join('; '), 'flags');
      flags.title = c.flags.join('\n');
      tr.appendChild(flags);

      var act = document.createElement('td');
      var del = document.createElement('button');
      del.className = 'mini danger';
      del.textContent = '×';
      del.title = 'Delete characteristic';
      del.onclick = function (e) {
        e.stopPropagation();
        state.chars = state.chars.filter(function (x) { return x.id !== c.id; });
        B.assignNumbers(state.chars);
        renderAll();
      };
      act.appendChild(del);
      tr.appendChild(act);

      tr.addEventListener('click', function () { select(c.id); });
      el.charBody.appendChild(tr);
    });
  }

  function cell(text, cls) {
    var td = document.createElement('td');
    if (cls) td.className = cls;
    td.textContent = text;
    return td;
  }

  function editCell(c, field, cls) {
    var td = document.createElement('td');
    if (cls) td.className = cls;
    var inp = document.createElement('input');
    inp.type = 'text';
    inp.value = c[field] == null ? '' : c[field];
    inp.addEventListener('change', function () {
      c[field] = inp.value;
      if (c.source === 'manual' && field === 'requirement') c.text = inp.value;
    });
    td.appendChild(inp);
    return td;
  }

  function numCell(c, field) {
    var td = document.createElement('td');
    td.className = 'numeric';
    var inp = document.createElement('input');
    inp.type = 'number';
    inp.step = 'any';
    inp.value = c[field] == null ? '' : c[field];
    inp.addEventListener('change', function () {
      c[field] = inp.value === '' ? null : parseFloat(inp.value);
    });
    td.appendChild(inp);
    return td;
  }

  function selectCell(c, field, options) {
    var td = document.createElement('td');
    var sel = document.createElement('select');
    options.forEach(function (o) {
      var opt = document.createElement('option');
      opt.value = o; opt.textContent = o;
      if (c[field] === o) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', function () { c[field] = sel.value; renderOverlay(); });
    td.appendChild(sel);
    return td;
  }

  function checkCell(c, field) {
    var td = document.createElement('td');
    td.className = 'center';
    var inp = document.createElement('input');
    inp.type = 'checkbox';
    inp.checked = !!c[field];
    inp.addEventListener('change', function () { c[field] = inp.checked; });
    td.appendChild(inp);
    return td;
  }

  function renderSummary() {
    var s = B.summarize(state.chars);
    var parts = ['<strong>' + s.total + '</strong> characteristics'];
    Object.keys(s.byType).sort().forEach(function (t) {
      parts.push('<span class="chip c-' + typeClass(t) + '">' + t + ': ' + s.byType[t] + '</span>');
    });
    if (s.review) parts.push('<span class="chip warn">Needs review: ' + s.review + '</span>');
    if (s.critical) parts.push('<span class="chip crit">Key: ' + s.critical + '</span>');
    el.summary.innerHTML = parts.join(' ');
  }

  // -------------------------------------------------------------------------
  // Navigation, persistence, export
  // -------------------------------------------------------------------------

  function gotoPage(i) {
    if (!state.pages.length) return;
    if (i < 0 || i >= state.pages.length) return;
    state.pageIndex = i;
    renderPage();
  }

  function setScale(s) {
    state.scale = Math.max(0.4, Math.min(5, s));
    renderPage();
  }

  function clearAll() {
    if (!confirm('Delete all ' + state.chars.length + ' characteristics? The loaded drawing stays open.')) return;
    state.chars = [];
    state.selectedId = null;
    renderAll();
    setStatus('Characteristic list cleared.');
  }

  function readMeta() {
    state.meta = {
      partNumber: el.partNumber.value,
      partName: el.partName.value,
      drawingNumber: el.drawingNumber.value,
      revision: el.revision.value,
      fairNumber: el.fairNumber.value,
      serialNumber: el.serialNumber.value,
      supplier: el.supplier.value,
      sourceFile: state.fileName
    };
    return state.meta;
  }

  function fileBase() {
    var pn = el.partNumber.value || state.fileName.replace(/\.pdf$/i, '') || 'drawing';
    return pn.replace(/[^\w.-]+/g, '_');
  }

  function saveLocal() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      meta: readMeta(), fileName: state.fileName, opts: state.opts, chars: state.chars
    }));
    setStatus('Saved ' + state.chars.length + ' characteristics to this browser.');
  }

  function loadLocal() {
    var raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) { setStatus('Nothing saved in this browser.', true); return; }
    try {
      var data = JSON.parse(raw);
      state.chars = data.chars || [];
      state.meta = data.meta || {};
      Object.keys(state.meta).forEach(function (k) { if (el[k]) el[k].value = state.meta[k] || ''; });
      renderAll();
      setStatus('Loaded ' + state.chars.length + ' saved characteristics' +
        (data.fileName && data.fileName !== state.fileName
          ? ' — saved against "' + data.fileName + '", which is not the file currently open.' : '.'));
    } catch (err) {
      setStatus('Saved data is corrupt: ' + err.message, true);
    }
  }

  /* The bible drawing itself: the rendered sheet with balloons and leaders burned
   * in, one PNG per sheet. */
  function exportBalloonedPng() {
    if (!state.pages.length) { setStatus('Load a PDF first.', true); return; }
    var src = el.pdfCanvas;
    var out = document.createElement('canvas');
    out.width = src.width; out.height = src.height;
    var ctx = out.getContext('2d');
    ctx.drawImage(src, 0, 0);

    var W = out.width, H = out.height;
    var r = Math.max(11, Math.round(W / 90));

    pageChars().forEach(function (c) {
      var bx = c.balloon.x * W, by = c.balloon.y * H;

      if (c.box) {
        ctx.strokeStyle = 'rgba(220,38,38,0.85)';
        ctx.lineWidth = 1.5;
        ctx.strokeRect(c.box.x * W - 2, c.box.y * H - 2, c.box.w * W + 4, c.box.h * H + 4);
        ctx.beginPath();
        ctx.moveTo(bx, by);
        ctx.lineTo((c.box.x + c.box.w / 2) * W, (c.box.y + c.box.h / 2) * H);
        ctx.stroke();
      }

      ctx.beginPath();
      ctx.arc(bx, by, r, 0, Math.PI * 2);
      ctx.fillStyle = c.critical ? '#fde68a' : '#ffffff';
      ctx.fill();
      ctx.lineWidth = 2;
      ctx.strokeStyle = '#dc2626';
      ctx.stroke();

      ctx.fillStyle = '#111827';
      ctx.font = 'bold ' + Math.round(r * 1.05) + 'px Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(String(c.number), bx, by + 1);
    });

    out.toBlob(function (blob) {
      downloadBlob(fileBase() + '-ballooned-sheet' + (state.pageIndex + 1) + '.png', blob);
    }, 'image/png');
  }

  function download(name, text, mime) {
    downloadBlob(name, new Blob([text], { type: mime + ';charset=utf-8;' }));
  }

  function downloadBlob(name, blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function setStatus(msg, isError) {
    el.status.textContent = msg;
    el.status.className = 'status' + (isError ? ' error' : '');
  }

  document.addEventListener('DOMContentLoaded', init);
})();
