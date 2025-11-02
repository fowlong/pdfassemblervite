// main.js
// --- light polyfill so pdfassembler’s CJS deps don’t choke in browsers
if (typeof window !== 'undefined') {
  window.global = window.global || window;
  window.process = window.process || {
    env: {},
    nextTick: (cb, ...args) => Promise.resolve().then(() => cb(...args))
  };
}

// ---- import pdfassembler on MAIN THREAD via shim (CJS-safe)
import PDFAssembler from './shims/pdfassembler.js';

// ---- pdf.js for canvas preview (kept separate from pdfassembler’s internal v2)
let pdfjsLib = null;
async function ensurePdfJs() {
  if (pdfjsLib) return pdfjsLib;
  pdfjsLib = await import('https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.mjs');
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.mjs';
  return pdfjsLib;
}

const $  = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

/* ---------------- State ---------------- */
let state = {
  asm: null,
  pdfTree: null,
  lastBlobUrl: null,
  lastUint8: null,
  pageCount: 0,
  assembling: false,
  inputFile: null
};

// --- FAST EDIT MODE knobs
const ASSEMBLE_DEBOUNCE_MS = 350;
const JSON_REFRESH_MS = 1200;
let assembleTimer = null;
let idleRehydrateTimer = null;
const dirtyPages = new Set();
let lastJsonRefresh = 0;

/* -------- Workers -------- */
let incremental = false;
let incWorker = null;
let incOpened = false;

let mapWorker = null;
function ensureMapperWorker() {
  if (mapWorker) return mapWorker;
  mapWorker = new Worker(new URL('./workers/mapper.worker.js', import.meta.url), {
    type: 'module',
    name: 'mapper'
  });
  return mapWorker;
}

function ensureIncWorker() {
  if (incWorker) return incWorker;
  incWorker = new Worker(new URL('./workers/incremental.worker.js', import.meta.url), {
    type: 'module',
    name: 'inc-writer'
  });
  return incWorker;
}
function rpc(worker, type, payload, transfer = []) {
  return new Promise((resolve, reject) => {
    const id = Math.random().toString(36).slice(2);
    const onMsg = (e) => {
      const d = e.data || {};
      if (d.id !== id) return;
      worker.removeEventListener('message', onMsg);
      if (d.ok) resolve(d.result);
      else reject(new Error(d.error || 'worker error'));
    };
    worker.addEventListener('message', onMsg);
    worker.postMessage({ id, type, payload }, transfer);
  });
}
function incCall(type, payload, transfer = []) {
  return rpc(ensureIncWorker(), type, payload, transfer);
}
function mapCall(type, payload, transfer = []) {
  return rpc(ensureMapperWorker(), type, payload, transfer);
}

/* ---------- ID helpers ---------- */
function objectIdOf(node) {
  if (node && node.$obj && Number.isFinite(node.$obj.n)) return { n: node.$obj.n|0, g: node.$obj.g|0 };
  if (node && node.__id && Number.isFinite(node.__id.n)) return { n: node.__id.n|0, g: node.__id.g|0 };
  if (node && node.$id  && Number.isFinite(node.$id.n))  return { n: node.$id.n|0,  g: node.$id.g|0  };
  if (node && node._id  && Number.isFinite(node._id.n))  return { n: node._id.n|0,  g: node._id.g|0  };
  if (node && node.$ref && Number.isFinite(node.$ref.n)) return { n: node.$ref.n|0,  g: node.$ref.g|0  };
  return null;
}

function flattenPagesFromTree(rootPages) {
  const out = [];
  function walk(node) {
    if (!node || typeof node !== 'object') return;
    const t = node['/Type'];
    const isPage  = (t === '/Page') || (t && String(t).includes('/Page')) ||
                    (node['/Contents'] != null && (node['/MediaBox'] || node['/CropBox']));
    const kids = node['/Kids'];

    if (isPage && !Array.isArray(kids)) {
      out.push(node);
      return;
    }
    if (Array.isArray(kids)) {
      for (const k of kids) walk(k);
    }
  }
  walk(rootPages);
  // if recursive walk found nothing, try direct /Kids fallback
  if (!out.length && Array.isArray(rootPages?.['/Kids'])) {
    for (const k of rootPages['/Kids']) walk(k);
  }
  return out;
}

function mappingFromTree(tree) {
  const rootPages = tree?.['/Root']?.['/Pages'];
  const pagesNodes = flattenPagesFromTree(rootPages);
  const pages = pagesNodes.map((p, i) => {
    const pageObj = objectIdOf(p) || null;
    const contents = [];
    const c = p?.['/Contents'];
    const arr = Array.isArray(c) ? c : (c ? [c] : []);
    arr.forEach((obj) => contents.push(objectIdOf(obj) || null));
    return { pageIndex: i, pageObj, contents };
  });

  return {
    pages,
    root: objectIdOf(tree?.['/Root']) || null,
    info: objectIdOf(tree?.['/Info']) || null
  };
}

// Merge tree mapping with byte-scan mapping so every page has concrete content IDs
async function mergeMappingWithBytes(bytesU8, map) {
  const copy = bytesU8.slice();
  // Scan ALL pages from bytes (no pre-known IDs required)
  const scan = await mapCall('scanAllPages', { bytes: copy.buffer }, [copy.buffer]);

  if (!scan || !Array.isArray(scan.pages) || !scan.pages.length) {
    // fallback: only resolve contents for pages where we DO know the page object id
    const needsHelp = map.pages.filter(pg => pg?.pageObj && (!pg.contents.length || pg.contents.some(x => !x)));
    if (needsHelp.length) {
      const buf2 = bytesU8.slice();
      const byPage = await mapCall(
        'contentsForPages',
        { bytes: buf2.buffer, pages: needsHelp.map(pg => pg.pageObj) },
        [buf2.buffer]
      );
      map.pages = map.pages.map(pg => {
        if (!pg?.pageObj) return pg;
        const key = `${pg.pageObj.n} ${pg.pageObj.g}`;
        const resolved = byPage[key];
        if (Array.isArray(resolved) && resolved.length) return { ...pg, contents: resolved };
        return pg;
      });
    }
    return map;
  }

  // Align lengths (prefer the scanned count if tree-derived was empty or mismatched)
  if (!map.pages.length || map.pages.length !== scan.pages.length) {
    map.pages = new Array(scan.pages.length).fill(0).map((_, i) => ({
      pageIndex: i, pageObj: null, contents: []
    }));
  }

  // Merge page-by-index: fill any missing pageObj/contents from scan
  map.pages = map.pages.map((pg, i) => {
    const s = scan.pages[i];
    const pageObj = pg.pageObj || s.pageObj || null;
    const contents =
      (!pg.contents.length || pg.contents.some(x => !x)) && Array.isArray(s.contents) && s.contents.length
        ? s.contents
        : pg.contents;
    return { pageIndex: i, pageObj, contents };
  });

  // Carry root/info if tree didn’t have them
  map.root = map.root || scan.root || null;
  map.info = map.info || scan.info || null;

  return map;
}

// NEW: open helpers
async function ensureIncOpenWith(bytesU8) {
  if (incOpened) return { pageCount: null };

  let mapping = mappingFromTree(state.pdfTree);
  mapping = await mergeMappingWithBytes(bytesU8, mapping);

  const allGood = mapping.pages.length &&
    mapping.pages.every(pg => Array.isArray(pg.contents) && pg.contents.length && pg.contents.every(Boolean));

  if (!allGood) {
    throw new Error('Incremental mapper could not resolve page/contents ids for this PDF.');
  }

  const copy = bytesU8.slice();
  const res = await incCall('open', { bytes: copy.buffer, mapping }, [copy.buffer]);
  incOpened = true;
  return res; // { pageCount }
}
async function ensureIncOpen() {
  if (incOpened) return;
  let base = state.lastUint8;
  if (!(base instanceof Uint8Array)) {
    base = await assembleOnceReturnBytes();
    state.lastUint8 = base;
  }
  await ensureIncOpenWith(base);
}

/* ---------------- Elements ---------------- */
const els = {
  fileInput: $('#fileInput'),
  loadSample: $('#loadSample'),
  assembleBtn: $('#assembleBtn'),
  downloadBtn: $('#downloadBtn'),
  pdfFrame: $('#pdfFrame'),
  jsonEditor: $('#jsonEditor'),
  refreshPreviewBtn: $('#refreshPreviewBtn'),
  syncFromTreeBtn: $('#syncFromTreeBtn'),
  indentToggle: $('#indentToggle'),
  compressToggle: $('#compressToggle'),
  removeRootBtn: $('#removeRootBtn'),
  status: $('#status'),
  canvasModeBtn: $('#canvasModeBtn'),
  tabs: $('#previewModeTabs'),
  panelNative: $('#panel-native'),
  panelCanvas: $('#panel-canvas'),
  canvasToolbar: $('#canvasToolbar'),
  addText: $('#addText'),
  commitOverlay: $('#commitOverlay'),
  pageIndex: $('#pageIndex'),
  renderPage: $('#renderPage'),
  pdfCanvas: $('#pdfCanvas'),
  overlayCanvas: $('#overlayCanvas'),
  qaFind: $('#qaFind'),
  qaReplace: $('#qaReplace'),
  qaRun: $('#qaRun'),
  // Text panel
  scanTextBtn: $('#scanTextBtn'),
  textFilter: $('#textFilter'),
  xTol: $('#xTol'),
  yTol: $('#yTol'),
  nudgeStep: $('#nudgeStep'),
  scanStatus: $('#scanStatus'),
  textGroups: $('#textGroups'),
  // XObject/Image panel
  scanXBtn: $('#scanXBtn'),
  xobjFilter: $('#xobjFilter'),
  xobjStatus: $('#xobjStatus'),
  xobjGroups: $('#xobjGroups'),
  xNudgeStep: $('#xNudgeStep'),
  xScaleStep: $('#xScaleStep'),
  // SVG/Paths/Tables
  scanPathsBtn: $('#scanPathsBtn'),
  pathFilter: $('#pathFilter'),
  pathStatus: $('#pathStatus'),
  pathGroups: $('#pathGroups'),
  pNudgeStep: $('#pNudgeStep'),
  pScaleStep: $('#pScaleStep'),
  // Safe panel
  scanSafeBtn: $('#scanSafeBtn'),
  safeFilter: $('#safeFilter'),
  safeStatus: $('#safeStatus'),
  safeGroups: $('#safeGroups'),
  // Incremental toggle
  incrementalToggle: $('#incrementalToggle'),
};

toast('Ready. Load a PDF to begin.');
wireUi();

/* ---------------- Scanner worker pool ---------------- */

class RPCWorker {
  constructor(url, { name = 'worker' } = {}) {
    this.worker = new Worker(url, { type: 'module', name });
    this._id = 1;
    this._pending = new Map();
    this.worker.onmessage = (e) => {
      const { id, ok, result, error } = e.data || {};
      if (!id) return;
      const p = this._pending.get(id);
      if (!p) return;
      this._pending.delete(id);
      if (ok) p.resolve(result);
      else p.reject(new Error(error || 'Worker error'));
    };
    this.worker.onerror = (err) => {
      for (const [, p] of this._pending) p.reject(err);
      this._pending.clear();
    };
  }
  call(type, payload, transfer = []) {
    const id = this._id++;
    const msg = { id, type, payload };
    const prom = new Promise((resolve, reject) => this._pending.set(id, { resolve, reject }));
    this.worker.postMessage(msg, transfer);
    return prom;
  }
  terminate() { this.worker.terminate(); }
}

// Tiny pool for scanner tasks (text + xobjects)
class ScannerPool {
  constructor(concurrency = 2) {
    this.workers = new Array(concurrency).fill(0).map((_, i) =>
      new RPCWorker(new URL('./workers/scanner.worker.js', import.meta.url), { name: `scanner-${i}` })
    );
    this.q = [];
    this.active = 0;
  }
  async _runOne(job) {
    this.active++;
    const w = this.workers[this.active % this.workers.length];
    try {
      const res = await w.call(job.type, job.payload);
      job.resolve(res);
    } catch (e) {
      job.reject(e);
    } finally {
      this.active--;
      this._pump();
    }
  }
  _pump() {
    while (this.active < this.workers.length && this.q.length) {
      const job = this.q.shift();
      this._runOne(job);
    }
  }
  submit(type, payload) {
    return new Promise((resolve, reject) => {
      this.q.push({ type, payload, resolve, reject });
      this._pump();
    });
  }
  terminate() { this.workers.forEach(w => w.terminate()); }
}
const scanners = new ScannerPool(2);

/* ------------- Vector worker (paths/rects) ------------- */
let vectorWorker = null;
function initVectorWorker() {
  try {
    vectorWorker = new Worker(new URL('./workers/vector.worker.js', import.meta.url), { type: 'module' });
  } catch (e) {
    console.warn('Vector worker unavailable; will fall back to sync scan.', e);
    vectorWorker = null;
  }
}
initVectorWorker();

/* ---------------- UI wiring ---------------- */

function wireUi() {
  els.tabs.addEventListener('click', (e) => {
    const btn = e.target.closest('button');
    if (!btn) return;
    $$('.tabs button').forEach(b => b.classList.toggle('active', b === btn));
    const tab = btn.dataset.tab;
    $$('.panel').forEach(p => p.classList.remove('active'));
    if (tab === 'native') els.panelNative.classList.add('active');
    if (tab === 'canvas') els.panelCanvas.classList.add('active');
  });

  els.fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (file) await loadPdf(file);
  });

  els.loadSample?.addEventListener('click', async () => {
    try {
      const res = await fetch('/public/sample.pdf');
      if (!res.ok) throw new Error('sample not found');
      const blob = await res.blob();
      await loadPdf(new File([blob], 'sample.pdf', { type: 'application/pdf' }));
    } catch {
      toast('No sample found at /public/sample.pdf. Use "Select PDF…".', true);
    }
  });

  els.refreshPreviewBtn.addEventListener('click', async () => {
    await applyJsonEditor(true);
  });

  els.syncFromTreeBtn.addEventListener('click', () => {
    if (!state.pdfTree) return;
    els.jsonEditor.value = stringifyPdfTree(state.pdfTree);
  });

  els.indentToggle.addEventListener('change', () => {});
  els.compressToggle.addEventListener('change', () => {});

  els.assembleBtn.addEventListener('click', () => forceAssembleAndFullRescan());

  els.removeRootBtn.addEventListener('click', async () => {
    if (!state.pdfTree) return;
    try {
      const root = state.pdfTree['/Root'];
      if (root) { delete root['/Outlines']; delete root['/PageLabels']; }
      maybeRefreshJsonEditor(true);
      scheduleAssemble({ reason: 'global' });
      toast('Removed /Outlines and /PageLabels from /Root (if present).');
    } catch (e) {
      console.error(e);
      toast('Failed removeRootEntries()', true);
    }
  });

  // Canvas mode
  els.canvasModeBtn.addEventListener('click', async () => {
    $$('.tabs button')[1].click();
    setupOverlayCanvases();
    els.canvasToolbar.hidden = false;
  });
  els.addText.addEventListener('click', () => addOverlayTextbox());
  els.renderPage.addEventListener('click', () => renderCanvasPage());
  els.commitOverlay.addEventListener('click', () => commitOverlayToPdf());

  // Quick action
  els.qaRun.addEventListener('click', () => runRegexReplace());

  // Text panel
  els.scanTextBtn?.addEventListener('click', () => { if (!state.pdfTree) return toast('Load a PDF first.', true); scanTextItems(); });
  els.textFilter?.addEventListener('input', () => renderTextGroups());
  els.xTol?.addEventListener('change', renderTextGroups);
  els.yTol?.addEventListener('change', renderTextGroups);

  // XObject/Image panel
  els.scanXBtn?.addEventListener('click', () => { if (!state.pdfTree) return toast('Load a PDF first.', true); scanXObjects(); });
  els.xobjFilter?.addEventListener('input', () => renderXGroups());

  // Paths panel
  els.scanPathsBtn?.addEventListener('click', () => { if (!state.pdfTree) return toast('Load a PDF first.', true); scanPaths(); });
  els.pathFilter?.addEventListener('input', () => renderPathGroups());
  els.pNudgeStep?.addEventListener('change', renderPathGroups);
  els.pScaleStep?.addEventListener('change', renderPathGroups);

  // Safe panel
  els.scanSafeBtn?.addEventListener('click', () => { if (!state.pdfTree) return toast('Load a PDF first.', true); scanSafeItems(); });
  els.safeFilter?.addEventListener('input', () => renderSafeGroups());

  // Incremental toggle (OPEN immediately)
  els.incrementalToggle?.addEventListener('change', async () => {
    if (els.incrementalToggle.checked) {
      try {
        await ensureIncOpen(); // opens WITH merged mapping
        const res = await incCall('getMapping', {});
        const pc = Array.isArray(res?.pages) ? res.pages.length : 0;
        if (!pc) throw new Error('Incremental mapper reported 0 pages.');
        incremental = true;
        toast('Incremental mode ON');
      } catch (e) {
        console.warn('Incremental open failed:', e);
        incremental = false;
        incOpened = false;
        els.incrementalToggle.checked = false;
        toast('Incremental mode unsupported for this PDF. Staying in full mode.', true);
      }
    } else {
      incremental = false;
      toast('Incremental mode OFF');
    }
  });
}

/* ---------------- Load / Assemble on main thread ---------------- */

async function loadPdf(file) {
  resetPreview();
  toast('Loading…');

  try {
    // start in fast mode
    els.compressToggle.checked = false;
    els.indentToggle.checked = false;

    // keep a handle to the input and its bytes
    state.inputFile = file;
    const buf = await file.arrayBuffer();
    const orig = new Uint8Array(buf);
    state.lastUint8 = orig;

    state.asm = new PDFAssembler(file);
    state.asm.compress = false;
    state.asm.indent = false;

    const pdfTree = await state.asm.getPDFStructure();
    state.pdfTree = pdfTree;

    try {
      state.pageCount = await state.asm.countPages();
      els.pageIndex.max = state.pageCount || 1;
    } catch {}

    maybeRefreshJsonEditor(true);
    els.assembleBtn.disabled = false;
    els.refreshPreviewBtn.disabled = false;
    els.syncFromTreeBtn.disabled = false;
    els.removeRootBtn.disabled = false;
    els.canvasModeBtn.disabled = false;

    setFrameUrl(URL.createObjectURL(file));
    toast('PDF loaded.');

    incOpened = false;

    if (els.incrementalToggle?.checked) {
      try {
        const res = await ensureIncOpenWith(orig);
        const pc = (res && typeof res.pageCount === 'number') ? res.pageCount : 0;
        if (!pc) throw new Error('Incremental mapper reported 0 pages.');
        console.log(`Incremental: mapping ready (${pc} page(s)).`);
        incremental = true;
      } catch (e) {
        console.warn('Incremental open failed; disabling incremental for this file.', e);
        incremental = false;
        incOpened = false;
        els.incrementalToggle.checked = false;
        toast('Incremental mode unsupported for this PDF (mapping failed). Using full mode.', true);
      }
    } else {
      incremental = false;
    }

    // initial scans (parallel)
    scanTextItems();
    scanXObjects();
    scanPaths();
    scanSafeItems();
  } catch (e) {
    console.error(e);
    toast('Error while loading PDF. See console.', true);
  }
}

function stringifyPdfTree(tree) {
  const seen = new WeakSet();
  return JSON.stringify(tree, function replacer(key, value) {
    if (typeof value === 'object' && value !== null) {
      if (seen.has(value)) return { $ref: true };
      seen.add(value);
    }
    return value;
  }, els.indentToggle.checked ? 2 : 0);
}

function maybeRefreshJsonEditor(force=false){
  const now = performance.now();
  if (!force && (now - lastJsonRefresh) < JSON_REFRESH_MS) return;
  lastJsonRefresh = now;
  requestIdleCallback?.(() => {
    els.jsonEditor.value = stringifyPdfTree(state.pdfTree);
  }, { timeout: 800 }) || (els.jsonEditor.value = stringifyPdfTree(state.pdfTree));
}

async function applyJsonEditor(andAssemble=false) {
  if (!state.pdfTree) return;
  try {
    const edited = JSON.parse(els.jsonEditor.value);
    state.pdfTree = edited;
    if (andAssemble) await forceAssembleAndFullRescan();
  } catch (e) {
    console.error(e);
    toast('JSON parse failed. Fix errors and try again.', true);
  }
}

function isNativeViewerActive() {
  const btn = document.querySelector('.tabs button[data-tab="native"]');
  return btn && btn.classList.contains('active');
}

async function assembleOnceReturnBytes() {
  if (!state.asm) return new Uint8Array();
  state.asm.compress = !!els.compressToggle.checked;
  state.asm.indent = els.indentToggle.checked ? 2 : false;
  state.asm.pdfTree = state.pdfTree;
  const uint8 = await state.asm.assemblePdf('Uint8Array');
  return uint8;
}

async function assembleAndPreviewConditional() {
  const bytes = await assembleOnceReturnBytes();
  state.lastUint8 = bytes;

  if (isNativeViewerActive()) {
    const blob = new Blob([bytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    setFrameUrl(url);
    els.downloadBtn.disabled = false;
    els.downloadBtn.onclick = () => downloadBlob(blob, 'edited.pdf');
  }
  toast('Assembled.');
}

async function rehydrateFromBytes(bytes) {
  const blob = new Blob([bytes], { type: 'application/pdf' });
  const asm = new PDFAssembler(blob);
  asm.indent = els.indentToggle.checked ? 2 : false;
  asm.compress = !!els.compressToggle.checked;
  const fresh = await asm.getPDFStructure();
  state.asm = asm;
  state.pdfTree = fresh;
  try {
    state.pageCount = await asm.countPages();
    els.pageIndex.max = state.pageCount || 1;
  } catch {}
}

async function forceAssembleAndFullRescan() {
  if (state.assembling) return;
  try {
    state.assembling = true;
    await assembleAndPreviewConditional();
    if (state.lastUint8) await rehydrateFromBytes(state.lastUint8);
    await Promise.all([ scanTextItems(), scanXObjects(), scanPaths(), scanSafeItems() ]);
    maybeRefreshJsonEditor(true);
  } catch (e) {
    console.error(e);
    toast('Assemble/rescan failed. See console.', true);
  } finally {
    state.assembling = false;
  }
}

// Debounced assemble + selective rescans
function scheduleAssemble({ pageIndex = null, reason = 'text' } = {}) {
  if (pageIndex != null) dirtyPages.add(pageIndex);
  if (assembleTimer) clearTimeout(assembleTimer);

  assembleTimer = setTimeout(async () => {
    assembleTimer = null;

    if (!incremental) {
      await assembleAndPreviewConditional();
    } else {
      try {
        await ensureIncOpen(); // mapping guaranteed
      } catch (e) {
        console.warn('Incremental ensure-open failed:', e);
        toast('Incremental mode failed to initialize. See console.', true);
        return;
      }

      const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'] || [];
      const pagesDirty = [...dirtyPages];
      const targetPages = pagesDirty.length ? pagesDirty : [];

      const edits = [];
      for (const p of targetPages) {
        const page = kids[p];
        if (!page) continue;
        const contents = page?.['/Contents'];
        const streams = Array.isArray(contents) ? contents : [contents];
        streams.forEach((obj, sIdx) => {
          if (!obj || typeof obj.stream !== 'string') return;
          const id = objectIdOf(obj);
          if (id) edits.push({ obj: id, streamText: obj.stream });
          else    edits.push({ pageIndex: p, streamIndex: sIdx, streamText: obj.stream });
        });
      }

      if (edits.length) {
        try {
          const res = await incCall('applyEdits', { edits });

          // Defensive copy of worker output so we never hold a transferred buffer
          const fromWorker = new Uint8Array(res.uint8);
          const bytes = new Uint8Array(fromWorker.length);
          bytes.set(fromWorker);

          state.lastUint8 = bytes;

          if (isNativeViewerActive()) {
            const blob = new Blob([bytes], { type: 'application/pdf' });
            const url = URL.createObjectURL(blob);
            setFrameUrl(url);
            els.downloadBtn.disabled = false;
            els.downloadBtn.onclick = () => downloadBlob(blob, 'edited.pdf');
          }
          toast('Incremental update appended.');
        } catch (e) {
          console.error(e);
          toast('Incremental update failed. See console.', true);
        }
      }
    }

    const pages = [...dirtyPages];
    dirtyPages.clear();
    if (pages.length) { scanTextItems(pages); scanXObjects(pages); scanPaths(pages); }
    else { scanTextItems(); scanXObjects(); scanPaths(); }

    if (reason !== 'text') scanSafeItems();

    if (idleRehydrateTimer) clearTimeout(idleRehydrateTimer);
    idleRehydrateTimer = setTimeout(async () => {
      if (!state.lastUint8) return;
      try {
        await rehydrateFromBytes(state.lastUint8);
        if (pages.length) { scanTextItems(pages); scanXObjects(pages); scanPaths(pages); }
        maybeRefreshJsonEditor(true);
      } catch (err) {
        console.warn('rehydrate failed (keeping previous tree):', err);
      }
    }, 1200);
  }, ASSEMBLE_DEBOUNCE_MS);
}

/* ---------------- Preview helpers ---------------- */

function setFrameUrl(url) {
  if (state.lastBlobUrl) URL.revokeObjectURL(state.lastBlobUrl);
  state.lastBlobUrl = url;
  els.pdfFrame.src = url;
}
function downloadBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
}
function toast(msg, error=false) {
  els.status.textContent = msg;
  els.status.style.color = error ? '#ffb0b0' : '#cce7ff';
  if (!error) console.log(msg);
}
function resetPreview(){
  if(state.lastBlobUrl) URL.revokeObjectURL(state.lastBlobUrl);
  state.lastBlobUrl=null;
  state.lastUint8=null;
  els.pdfFrame.removeAttribute('src');
}

/* ---------------- Quick Action: regex replace ---------------- */

function runRegexReplace() {
  if (!state.pdfTree) return;
  let find = els.qaFind.value ?? '';
  let repl = els.qaReplace.value ?? '';
  if (!find.length) return toast('Enter a regex to find.', true);

  function buildRegExp(input, defaultFlags='g') {
    const s = String(input).trim();
    if (s.startsWith('/') && s.lastIndexOf('/') > 0) {
      const last = s.lastIndexOf('/');
      const pat = s.slice(1, last);
      const flags = s.slice(last + 1) || defaultFlags;
      return new RegExp(pat, flags);
    }
    return new RegExp(s, defaultFlags);
  }

  try {
    const re = buildRegExp(find, 'g');
    const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'];
    if (!Array.isArray(kids)) return toast('No /Kids array found in /Root./Pages', true);
    for (const page of kids) {
      const contents = page?.['/Contents']; if (!contents) continue;
      const streams = Array.isArray(contents) ? contents : [contents];
      for (const s of streams) {
        if (!s || typeof s !== 'object') continue;
        if (typeof s.stream === 'string') {
          try { s.stream = s.stream.replace(re, repl); }
          catch (inner) { console.error('Replace failed on a stream:', inner); }
        }
      }
    }
    maybeRefreshJsonEditor(true);
    scheduleAssemble({ reason: 'global' });
  } catch (e) { console.error(e); toast('Regex failed: ' + (e?.message || 'invalid pattern'), true); }
}

/* ---------------- Canvas (preview) ---------------- */

function setupOverlayCanvases() {
  const wrap = document.querySelector('.canvasWrap');
  const rect = wrap.getBoundingClientRect();
  [els.pdfCanvas, els.overlayCanvas].forEach(c => { c.width = Math.max(800, Math.floor(rect.width - 2)); c.height = 1000; });
  const ctx = els.overlayCanvas.getContext('2d'); ctx.lineWidth = 1; ctx.strokeStyle = '#66ccff';
}
async function renderCanvasPage() {
  if (!state.lastUint8) scheduleAssemble({});
  await ensurePdfJs();
  const data = state.lastUint8;
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const index1 = Math.min(Math.max(1, parseInt(els.pageIndex.value || '1', 10)), pdf.numPages);
  els.pageIndex.value = index1;
  const page = await pdf.getPage(index1);
  const viewport = page.getViewport({ scale: 1.5 });
  const canvas = els.pdfCanvas;
  canvas.width = viewport.width; canvas.height = viewport.height;
  const ctx = canvas.getContext('2d');
  await page.render({ canvasContext: ctx, viewport }).promise;
  els.overlayCanvas.width = canvas.width; els.overlayCanvas.height = canvas.height;
  clearOverlay();
}
function clearOverlay() {
  const octx = els.overlayCanvas.getContext('2d');
  octx.clearRect(0,0,els.overlayCanvas.width, els.overlayCanvas.height);
  overlayItems.length = 0;
}

// Simple draggable overlay
const overlayItems = []; let dragIndex = -1; let lastDrag=null;
function addOverlayTextbox(){ overlayItems.push({ x:100, y:100, text:'New text' }); drawOverlay(); }
function drawOverlay(){
  const ctx=els.overlayCanvas.getContext('2d');
  ctx.clearRect(0,0,els.overlayCanvas.width,els.overlayCanvas.height);
  ctx.font='16px sans-serif'; ctx.fillStyle='#ffffff'; ctx.strokeStyle='#66ccff';
  overlayItems.forEach(it=>{
    ctx.fillText(it.text,it.x,it.y);
    const w=ctx.measureText(it.text).width+8, h=22;
    ctx.strokeRect(it.x-4,it.y-16,w,h);
  });
}
els.overlayCanvas.addEventListener('mousedown',e=>{ const {x,y}=overlayPos(e); const i=hitIndex(x,y); dragIndex=i; lastDrag={x,y}; });
els.overlayCanvas.addEventListener('mousemove',e=>{ if(dragIndex<0) return; const {x,y}=overlayPos(e); const dx=x-lastDrag.x,dy=y-lastDrag.y; overlayItems[dragIndex].x+=dx; overlayItems[dragIndex].y+=dy; lastDrag={x,y}; drawOverlay(); });
['mouseup','mouseleave'].forEach(ev=>els.overlayCanvas.addEventListener(ev,()=>dragIndex=-1));
els.overlayCanvas.addEventListener('dblclick',e=>{ const {x,y}=overlayPos(e); const i=hitIndex(x,y); if(i>=0){ const t=prompt('Edit text:',overlayItems[i].text); if(t!=null){overlayItems[i].text=t; drawOverlay();}}});
function overlayPos(e){ const r=els.overlayCanvas.getBoundingClientRect(); return {x:e.clientX-r.left,y:e.clientY-r.top}; }
function hitIndex(x,y){ const ctx=els.overlayCanvas.getContext('2d'); ctx.font='16px sans-serif'; for(let i=overlayItems.length-1;i>=0;i--){ const it=overlayItems[i]; const w=ctx.measureText(it.text).width+8,h=22; if(x>=it.x-4&&x<=it.x-4+w&&y>=it.y-16&&y<=it.y-16+h) return i;} return -1; }

/* ---------------- Text scanning & editing (scanner workers) ---------------- */

let TEXT_GROUPS = []; // [{pageIndex, items:[...]}]

async function scanTextItems(onlyPages = null){
  if (!state.pdfTree) return;
  const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'];
  const allPages = Array.isArray(kids) ? kids.map((_,i)=>i) : [];
  const pages = Array.isArray(onlyPages) && onlyPages.length ? onlyPages : allPages;

  const treeJSON = stringifyPdfTree(state.pdfTree);
  const xTol = Number(els.xTol?.value || 4);
  const yTol = Number(els.yTol?.value || 2);

  const jobs = pages.map(p => scanners.submit('scanText', { treeJSON, pageIndex: p, xTol, yTol }));
  const results = await Promise.all(jobs);

  const maxPage = (kids?.length || 0) - 1;
  const pageMap = new Map(results.map(r => [r.pageIndex, r]));
  TEXT_GROUPS = [];
  for (let i=0;i<=maxPage;i++){
    TEXT_GROUPS.push(pageMap.get(i) || { pageIndex:i, items:[] });
  }
  renderTextGroups();
}

function renderTextGroups(){
  const wrap = els.textGroups; if (!wrap) return;
  const filter = (els.textFilter?.value || '').toLowerCase();
  const step   = Number(els.nudgeStep?.value || 1);
  wrap.innerHTML = '';

  let total = 0;
  TEXT_GROUPS.forEach(group=>{
    const visible = group.items.filter(it => !filter || String(it.text).toLowerCase().includes(filter));
    total += visible.length;
    const box = document.createElement('div');
    box.className = 'text-group';
    box.innerHTML = `<h3>Page ${group.pageIndex+1} — ${visible.length} item(s)</h3>`;
    visible.forEach(it => box.appendChild(renderTextItem(it, step)));
    wrap.appendChild(box);
  });
  if (els.scanStatus) els.scanStatus.textContent = `${total} text item(s).`;
}

function renderTextItem(it, step){
  const d = document.createElement('details');
  d.className = 'text-item';
  const preview = it.text.replace(/\s+/g,' ').trim();
  d.innerHTML = `
    <summary>
      <span class="ellipsis">${escapeHtml(preview || '␣')}</span>
      <span class="meta">
        <span class="badge">p${it.pageIndex+1}</span>
        <span class="badge">${it.posType || '?'}</span>
        <span class="badge">x:${fmtNum(it.x)} y:${fmtNum(it.y)}</span>
        ${Number.isFinite(it.fontSize)? `<span class="badge">size:${fmtNum(it.fontSize)}</span>`:''}
      </span>
    </summary>
    <div class="edit">
      <input class="ti-text" type="text" value="${escapeAttr(it.text)}" />
      <div class="text-xy">
        <label>X: <input class="ti-x" type="number" step="0.5" value="${numOrEmpty(it.x)}"></label>
        <label>Y: <input class="ti-y" type="number" step="0.5" value="${numOrEmpty(it.y)}"></label>
        <span style="opacity:.7;font-size:12px">(Tm/Td)</span>
      </div>
      <div class="font-mini">
        <label>Size: <input class="ti-fs" type="number" step="0.5" value="${Number.isFinite(it.fontSize)? it.fontSize : ''}"></label>
        <button class="ti-applyfs" title="Apply font size">Set</button>
      </div>
      <div class="nudges">
        <button class="nL">←</button>
        <button class="nU">↑</button>
        <button class="nD">↓</button>
        <button class="nR">→</button>
        <button class="ti-apply">Apply text</button>
      </div>
    </div>
  `;

  const q = sel => d.querySelector(sel);
  q('.nL').onclick = async () => { await nudgeItem(it, -step, 0); };
  q('.nR').onclick = async () => { await nudgeItem(it, +step, 0); };
  q('.nU').onclick = async () => { await nudgeItem(it, 0, +step); };
  q('.nD').onclick = async () => { await nudgeItem(it, 0, -step); };
  q('.ti-apply').onclick = async () => { await replaceGroupText(it, q('.ti-text').value ?? ''); };
  q('.ti-x').onchange = async () => { await setItemXY(it, Number(q('.ti-x').value), it.y); };
  q('.ti-y').onchange = async () => { await setItemXY(it, it.x, Number(q('.ti-y').value)); };
  q('.ti-applyfs').onclick = async () => {
    const fs = Number(q('.ti-fs').value);
    if (!Number.isFinite(fs) || fs <= 0) return toast('Enter a valid font size', true);
    await applyFontSize(it, fs);
  };

  return d;
}

function getObjRefForText(it){
  const page = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids']?.[it.pageIndex];
  const contents = page?.['/Contents'];
  const streams = Array.isArray(contents) ? contents : [contents];
  const obj = streams[it.streamIndex];
  return (obj && typeof obj.stream === 'string') ? obj : null;
}

async function replaceGroupText(it, newText){
  const obj = getObjRefForText(it); if(!obj) return toast('Stream not found.', true);
  const s = obj.stream;
  const replacement = `[(${pdfEscapeLiteral(newText)})] TJ`;

  const slice = s.slice(it.start, it.end);
  if (slice.length && s.indexOf(slice, Math.max(0, it.start - 4)) >= 0) {
    obj.stream = s.slice(0, it.start) + replacement + s.slice(it.end);
  } else {
    const local = localReFindStreamToken(s, it.start, it.end);
    if (!local) return toast('Could not locate token to replace.', true);
    obj.stream = s.slice(0, local.a) + replacement + s.slice(local.b);
  }

  it.text = newText;
  maybeRefreshJsonEditor(true);
  scheduleAssemble({ pageIndex: it.pageIndex, reason: 'text' });
}

function localReFindStreamToken(s, start, end){
  const isWs = ch => ch===' '||ch==='\n'||ch==='\r'||ch==='\t'||ch==='\f'||ch=='\v';
  const winStart = Math.max(0, start - 400);
  const winEnd   = Math.min(s.length, end + 400);
  const win = s.slice(winStart, winEnd);
  const relStart = start - winStart;
  const i = win.lastIndexOf('[', relStart);
  let j = -1;
  if (i >= 0) {
    const close = win.indexOf(']', i + 1);
    if (close > i) {
      let k = close + 1; while (k < win.length && isWs(win[k])) k++;
      if (win.slice(k, k + 2) === 'TJ') j = k + 2;
    }
  }
  if (i >= 0 && j > i) return { a: winStart + i, b: winStart + j };
  const k = win.lastIndexOf('(', relStart);
  const l = win.indexOf(') Tj', Math.max(k + 1, 0));
  if (k >= 0 && l > k) return { a: winStart + k, b: winStart + l + 3 };
  return null;
}

async function nudgeItem(it, dx, dy){
  if (!it.posType || it.posIndex == null) return toast('No position operator near this text; cannot nudge.', true);
  const obj = getObjRefForText(it); if(!obj) return toast('Stream not found.', true);
  const s = obj.stream;

  if (it.posType === 'Td') {
    const m = /(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+Td/.exec(s.slice(it.posIndex));
    if (!m) return toast('Td not found.', true);
    const nx = Number(m[1]) + dx, ny = Number(m[2]) + dy;
    obj.stream = s.slice(0, it.posIndex) + `${nx} ${ny} Td` + s.slice(it.posIndex + m[0].length);
    it.x = nx; it.y = ny;
  } else {
    const m = /(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+Tm/.exec(s.slice(it.posIndex));
    if (!m) return toast('Tm not found.', true);
    const e = Number(m[5]) + dx, f = Number(m[6]) + dy;
    obj.stream = s.slice(0, it.posIndex) + `${m[1]} ${m[2]} ${m[3]} ${m[4]} ${e} ${f} Tm` + s.slice(it.posIndex + m[0].length);
    it.x = e; it.y = f;
  }
  maybeRefreshJsonEditor(true);
  scheduleAssemble({ pageIndex: it.pageIndex, reason: 'text' });
}
async function setItemXY(it, x, y){ const dx = Number.isFinite(x)&&Number.isFinite(it.x)?(x-it.x):0; const dy = Number.isFinite(y)&&Number.isFinite(it.y)?(y-it.y):0; if (!dx && !dy) return; await nudgeItem(it, dx, dy); }

async function applyFontSize(it, size){
  const obj = getObjRefForText(it); if(!obj) return toast('Stream not found.', true);
  const s = obj.stream;
  const tf = lastMatchBefore(/\/([^\s]+)\s+(-?\d+(?:\.\d+)?)\s+Tf/g, s, it.start);
  if (tf) {
    obj.stream = s.slice(0, tf.index) + `/${tf[1]} ${size} Tf` + s.slice(tf.index + tf[0].length);
    it.fontSize = size;
  } else {
    obj.stream = s.slice(0, it.start) + `/Helv ${size} Tf ` + s.slice(it.start);
    it.fontSize = size;
  }
  maybeRefreshJsonEditor(true);
  scheduleAssemble({ pageIndex: it.pageIndex, reason: 'text' });
}

/* ---------------- XObjects / Images ---------------- */

let XOBJ_GROUPS = []; // [{pageIndex, items:[...]}]

async function scanXObjects(onlyPages = null){
  if (!state.pdfTree) return;
  const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'];
  const allPages = Array.isArray(kids) ? kids.map((_,i)=>i) : [];
  const pages = Array.isArray(onlyPages) && onlyPages.length ? onlyPages : allPages;

  const treeJSON = stringifyPdfTree(state.pdfTree);
  const jobs = pages.map(p => scanners.submit('scanXObjects', { treeJSON, pageIndex: p }));
  const results = await Promise.all(jobs);

  const maxPage = (kids?.length || 0) - 1;
  const pageMap = new Map(results.map(r => [r.pageIndex, r]));
  XOBJ_GROUPS = [];
  for (let i=0;i<=maxPage;i++){
    XOBJ_GROUPS.push(pageMap.get(i) || { pageIndex:i, items:[] });
  }
  renderXGroups();
}

function renderXGroups(){
  const wrap = els.xobjGroups; if (!wrap) return;
  const filter = (els.xobjFilter?.value || '').toLowerCase();
  const step   = Number(els.xNudgeStep?.value || 1);
  const sStep  = Number(els.xScaleStep?.value || 0.05);
  wrap.innerHTML = '';

  let total = 0;
  XOBJ_GROUPS.forEach(group=>{
    const visible = group.items.filter(it => {
      const label = `${it.name} ${it.kind}`.toLowerCase();
      return !filter || label.includes(filter);
    });
    total += visible.length;
    const box = document.createElement('div');
    box.className = 'text-group';
    box.innerHTML = `<h3>Page ${group.pageIndex+1} — ${visible.length} XObject draw(s)</h3>`;
    visible.forEach(it => box.appendChild(renderXItem(it, step, sStep)));
    wrap.appendChild(box);
  });
  if (els.xobjStatus) els.xobjStatus.textContent = `${total} XObject draw(s).`;
}

function renderXItem(it, step, sStep){
  const d = document.createElement('details');
  d.className = 'text-item';
  const axisAligned = Math.abs(it.b) < 1e-6 && Math.abs(it.c) < 1e-6;
  d.innerHTML = `
    <summary>
      <span class="ellipsis">/${escapeHtml(it.name)} <small>(${it.kind})</small></span>
      <span class="meta">
        <span class="badge">p${it.pageIndex+1}</span>
        <span class="badge">${it.hasCM ? 'cm' : 'no cm'}</span>
        <span class="badge">x:${fmtNum(it.e)} y:${fmtNum(it.f)}</span>
        ${axisAligned ? `<span class="badge">sx:${fmtNum(it.a)} sy:${fmtNum(it.d)}</span>` : `<span class="badge">matrix</span>`}
      </span>
    </summary>
    <div class="edit">
      <div class="text-xy">
        <label>X: <input class="xo-x" type="number" step="0.5" value="${numOrEmpty(it.e)}"></label>
        <label>Y: <input class="xo-y" type="number" step="0.5" value="${numOrEmpty(it.f)}"></label>
        <span style="opacity:.7;font-size:12px">(cm tx/ty)</span>
      </div>
      <div class="font-mini">
        <label>sX: <input class="xo-sx" type="number" step="0.01" value="${it.a}"></label>
        <label>sY: <input class="xo-sy" type="number" step="0.01" value="${it.d}"></label>
        <button class="xo-apply" title="Apply matrix">Apply</button>
      </div>
      <div class="nudges">
        <button class="nL">←</button>
        <button class="nU">↑</button>
        <button class="nD">↓</button>
        <button class="nR">→</button>
        <button class="sDown">– size</button>
        <button class="sUp">+ size</button>
        <button class="reset">reset cm</button>
      </div>
    </div>
  `;

  const q = sel => d.querySelector(sel);
  q('.nL').onclick = async () => { await xNudge(it, -step, 0); };
  q('.nR').onclick = async () => { await xNudge(it, +step, 0); };
  q('.nU').onclick = async () => { await xNudge(it, 0, +step); };
  q('.nD').onclick = async () => { await xNudge(it, 0, -step); };
  q('.sUp').onclick = async () => { await xScale(it, 1+sStep, 1+sStep); };
  q('.sDown').onclick = async () => { await xScale(it, 1/(1+sStep), 1/(1+sStep)); };
  q('.reset').onclick = async () => { await xApplyMatrix(it, 1,0,0,1,0,0); };

  q('.xo-apply').onclick = async () => {
    const sx = Number(q('.xo-sx').value); const sy = Number(q('.xo-sy').value);
    const tx = Number(q('.xo-x').value);  const ty = Number(q('.xo-y').value);
    if (![sx,sy,tx,ty].every(Number.isFinite)) return toast('Enter valid numbers', true);
    await xApplyMatrix(it, sx, 0, 0, sy, tx, ty);
  };
  q('.xo-x').onchange = async () => { const tx = Number(q('.xo-x').value); if(Number.isFinite(tx)) await xApplyMatrix(it, it.a,it.b,it.c,it.d,tx,it.f); };
  q('.xo-y').onchange = async () => { const ty = Number(q('.xo-y').value); if(Number.isFinite(ty)) await xApplyMatrix(it, it.a,it.b,it.c,it.d,it.e,ty); };

  return d;
}

function getStreamForX(it){
  const page = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids']?.[it.pageIndex];
  const contents = page?.['/Contents']; const streams = Array.isArray(contents) ? contents : [contents];
  const obj = streams[it.streamIndex];
  return (obj && typeof obj.stream === 'string') ? obj : null;
}
function fmt(n){ return (Math.round(n*10000)/10000).toString(); }

async function xApplyMatrix(it, a,b,c,d,e,f){
  const obj = getStreamForX(it); if(!obj) return toast('Stream not found.', true);
  const s = obj.stream;
  const token = `${fmt(a)} ${fmt(b)} ${fmt(c)} ${fmt(d)} ${fmt(e)} ${fmt(f)} cm`;

  if (it.hasCM && it.cmIndex != null && typeof it.cmText === 'string') {
    obj.stream = s.slice(0, it.cmIndex) + token + s.slice(it.cmIndex + it.cmText.length);
  } else {
    obj.stream = s.slice(0, it.doStart) + token + ' ' + s.slice(it.doStart);
  }

  it.hasCM = true; it.cmText = token; it.a=a; it.b=b; it.c=c; it.d=d; it.e=e; it.f=f;
  maybeRefreshJsonEditor(true);
  scheduleAssemble({ pageIndex: it.pageIndex, reason: 'text' });
}
async function xNudge(it, dx, dy){ await xApplyMatrix(it, it.a, it.b, it.c, it.d, it.e + dx, it.f + dy); }
async function xScale(it, sx, sy){ await xApplyMatrix(it, it.a * sx, it.b, it.c, it.d * sy, it.e, it.f); }

/* ---------------- Paths / SVG-ish (rects + paint ops) ---------------- */

let PATH_GROUPS = []; // [{pageIndex, items:[{kind:'rect'|'path', op, streamIndex, start,end, hasCM, cmIndex, cmText, a,b,c,d,e,f, rect?}]}]

function renderPathGroups(){
  const wrap = els.pathGroups; if (!wrap) return;
  const filter = (els.pathFilter?.value || '').toLowerCase();
  const step   = Number(els.pNudgeStep?.value || 1);
  const sStep  = Number(els.pScaleStep?.value || 0.05);
  wrap.innerHTML = '';

  let total = 0;
  PATH_GROUPS.forEach(group=>{
    const visible = group.items.filter(it => !filter || (`${it.kind} ${it.op||''}`).toLowerCase().includes(filter));
    total += visible.length;
    const box = document.createElement('div'); box.className='text-group';
    box.innerHTML = `<h3>Page ${group.pageIndex+1} — ${visible.length} vector item(s)</h3>`;
    visible.forEach(it => box.appendChild(renderPathItem(group.pageIndex, it, step, sStep)));
    wrap.appendChild(box);
  });
  if (els.pathStatus) els.pathStatus.textContent = `${total} vector item(s).`;
}

function renderPathItem(pageIndex, it, step, sStep){
  const d = document.createElement('details'); d.className='text-item';
  const axisAligned = Math.abs(it.b||0) < 1e-6 && Math.abs(it.c||0) < 1e-6;
  d.innerHTML = `
    <summary>
      <span class="ellipsis">${it.kind} <small>${it.op || ''}</small></span>
      <span class="meta">
        <span class="badge">p${pageIndex+1}</span>
        <span class="badge">${it.hasCM ? 'cm' : 'no cm'}</span>
        <span class="badge">x:${fmtNum(it.e||0)} y:${fmtNum(it.f||0)}</span>
        ${axisAligned ? `<span class="badge">sx:${fmtNum(it.a||1)} sy:${fmtNum(it.d||1)}</span>` : `<span class="badge">matrix</span>`}
      </span>
    </summary>
    <div class="edit">
      <div class="text-xy">
        <label>X: <input class="vp-x" type="number" step="0.5" value="${numOrEmpty(it.e)}"></label>
        <label>Y: <input class="vp-y" type="number" step="0.5" value="${numOrEmpty(it.f)}"></label>
      </div>
      <div class="font-mini">
        <label>sX: <input class="vp-sx" type="number" step="0.01" value="${it.a ?? 1}"></label>
        <label>sY: <input class="vp-sy" type="number" step="0.01" value="${it.d ?? 1}"></label>
        <button class="vp-apply">Apply</button>
      </div>
      <div class="nudges">
        <button class="nL">←</button>
        <button class="nU">↑</button>
        <button class="nD">↓</button>
        <button class="nR">→</button>
        <button class="sDown">– size</button>
        <button class="sUp">+ size</button>
        <button class="reset">reset cm</button>
      </div>
    </div>
  `;
  const q = s => d.querySelector(s);
  q('.nL').onclick = async () => pNudge(pageIndex, it, -step, 0);
  q('.nR').onclick = async () => pNudge(pageIndex, it, +step, 0);
  q('.nU').onclick = async () => pNudge(pageIndex, it, 0, +step);
  q('.nD').onclick = async () => pNudge(pageIndex, it, 0, -step);
  q('.sUp').onclick = async () => pScale(pageIndex, it, 1+sStep, 1+sStep);
  q('.sDown').onclick = async () => pScale(pageIndex, it, 1/(1+sStep), 1/(1+sStep));
  q('.reset').onclick = async () => pApplyMatrix(pageIndex, it, 1,0,0,1,0,0);
  q('.vp-apply').onclick = async () => {
    const sx = Number(q('.vp-sx').value), sy = Number(q('.vp-sy').value);
    const tx = Number(q('.vp-x').value),  ty = Number(q('.vp-y').value);
    if (![sx,sy,tx,ty].every(Number.isFinite)) return toast('Enter valid numbers', true);
    await pApplyMatrix(pageIndex, it, sx,0,0,sy,tx,ty);
  };
  q('.vp-x').onchange = async () => { const tx = Number(q('.vp-x').value); if(Number.isFinite(tx)) await pApplyMatrix(pageIndex, it, it.a||1,it.b||0,it.c||0,it.d||1,tx,it.f||0); };
  q('.vp-y').onchange = async () => { const ty = Number(q('.vp-y').value); if(Number.isFinite(ty)) await pApplyMatrix(pageIndex, it, it.a||1,it.b||0,it.c||0,it.d||1,it.e||0,ty); };
  return d;
}

function getStreamBy(pageIndex, streamIndex){
  const page = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids']?.[pageIndex];
  const contents = page?.['/Contents']; const streams = Array.isArray(contents) ? contents : [contents];
  const obj = streams[streamIndex];
  return (obj && typeof obj.stream === 'string') ? obj : null;
}
function pfmt(n){ return (Math.round(n*10000)/10000).toString(); }

async function pApplyMatrix(pageIndex, it, a,b,c,d,e,f){
  const obj = getStreamBy(pageIndex, it.streamIndex); if(!obj) return toast('Stream not found.', true);
  const s = obj.stream;
  const token = `${pfmt(a)} ${pfmt(b)} ${pfmt(c)} ${pfmt(d)} ${pfmt(e)} ${pfmt(f)} cm`;
  if (it.hasCM && it.cmIndex != null && typeof it.cmText === 'string') {
    obj.stream = s.slice(0, it.cmIndex) + token + s.slice(it.cmIndex + it.cmText.length);
  } else {
    obj.stream = s.slice(0, it.start) + token + ' ' + s.slice(it.start);
  }
  it.hasCM = true; it.cmText = token; it.a=a; it.b=b; it.c=c; it.d=d; it.e=e; it.f=f;
  maybeRefreshJsonEditor(true);
  scheduleAssemble({ pageIndex, reason: 'vector' });
}
async function pNudge(pageIndex, it, dx, dy){ await pApplyMatrix(pageIndex, it, it.a||1, it.b||0, it.c||0, it.d||1, (it.e||0)+dx, (it.f||0)+dy); }
async function pScale(pageIndex, it, sx, sy){ await pApplyMatrix(pageIndex, it, (it.a||1)*sx, it.b||0, it.c||0, (it.d||1)*sy, it.e||0, it.f||0); }

/* Vector scanning with worker (or sync fallback) */
function scanPaths(onlyPages = null){
  if (!state.pdfTree) return;
  const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'];
  const allPages = Array.isArray(kids) ? kids.map((_,i)=>i) : [];
  const pages = Array.isArray(onlyPages) && onlyPages.length ? onlyPages : allPages;

  const keep = new Map();
  if (onlyPages && Array.isArray(PATH_GROUPS) && PATH_GROUPS.length) {
    for (const g of PATH_GROUPS) keep.set(g.pageIndex, { pageIndex:g.pageIndex, items:[...g.items] });
  }

  const payloadPages = pages.map(pIndex => {
    const page = kids?.[pIndex];
    const contents = page?.['/Contents']; const streams = Array.isArray(contents)? contents : [contents];
    const streamText = streams.map(o => (o && typeof o.stream === 'string') ? o.stream : '');
    return { pageIndex: pIndex, streams: streamText };
  });

  if (vectorWorker) {
    const onMessage = (ev) => {
      const { type, results } = ev.data || {};
      if (type !== 'scanPaths:done') return;
      vectorWorker.removeEventListener('message', onMessage);
      results.forEach(pg => {
        keep.set(pg.pageIndex, { pageIndex: pg.pageIndex, items: pg.items });
      });
      rebuildPathGroups(keep, kids?.length||0);
      renderPathGroups();
    };
    vectorWorker.addEventListener('message', onMessage);
    vectorWorker.postMessage({ type:'scanPaths', payload: { pages: payloadPages, options: {} } });
  } else {
    for (const p of payloadPages) {
      const items = [];
      p.streams.forEach((src, streamIndex) => {
        if (typeof src !== 'string' || !src.length) return;
        const cms=[]; let m;
        const cmRe=/(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+cm/g;
        while((m=cmRe.exec(src))) cms.push({ index:m.index, text:m[0], a:+m[1], b:+m[2], c:+m[3], d:+m[4], e:+m[5], f:+m[6] });
        const lastCmBefore = pos => { let last=null; for (let i=0;i<cms.length;i++){ if (cms[i].index<pos) last=cms[i]; else break; } return last; };
        const reRect=/(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+re/g;
        while((m=reRect.exec(src))){
          const near=lastCmBefore(m.index);
          items.push({ kind:'rect', op:'re', streamIndex, start:m.index, end:reRect.lastIndex,
            hasCM:!!near, cmIndex:near?.index??null, cmText:near?.text??null,
            a:near?.a??1,b:near?.b??0,c:near?.c??0,d:near?.d??1,e:near?.e??0,f:near?.f??0,
            rect:{x:+m[1],y:+m[2],w:+m[3],h:+m[4]} });
        }
        const paintRe = /\b(S|s|f\*?|B\*?|b\*?|n)\b/g;
        while((m=paintRe.exec(src))){
          const near=lastCmBefore(m.index);
          items.push({ kind:'path', op:m[1], streamIndex, start:m.index, end:paintRe.lastIndex,
            hasCM:!!near, cmIndex:near?.index??null, cmText:near?.text??null,
            a:near?.a??1,b:near?.b??0,c:near?.c??0,d:near?.d??1,e:near?.e??0,f:near?.f??0 });
        }
      });
      keep.set(p.pageIndex, { pageIndex:p.pageIndex, items });
    }
    rebuildPathGroups(keep, kids?.length||0);
    renderPathGroups();
  }
}

function rebuildPathGroups(keep, totalPages){
  PATH_GROUPS = [];
  for (let i=0;i<totalPages;i++){
    PATH_GROUPS.push(keep.get(i) || { pageIndex:i, items:[] });
  }
}

/* ---------------- Safe items ---------------- */

let SAFE_GROUPS = [];

function scanSafeItems(){
  SAFE_GROUPS = [];

  const info = state.pdfTree?.['/Info'];
  if (info && typeof info === 'object') {
    const items = [];
    for (const k of Object.keys(info)) {
      if (!k.startsWith('/')) continue;
      const v = info[k];
      if (['string','number','boolean'].includes(typeof v)) {
        items.push(safeItem('info', `${k}`, () => info[k], (nv)=>{ info[k] = nv; }));
      }
    }
    if (items.length) SAFE_GROUPS.push({ title: 'Document Info', items });
  }

  const fields = state.pdfTree?.['/Root']?.['/AcroForm']?.['/Fields'];
  if (Array.isArray(fields)) {
    const items = [];
    walkFields(fields, (dict, path) => {
      const name = dict['/T'];
      const val  = dict['/V'];
      if (typeof name === 'string' && typeof val === 'string') {
        items.push(safeItem('form', `${path} (${name})`, () => dict['/V'], (nv)=>{ dict['/V'] = nv; }));
      }
    });
    if (items.length) SAFE_GROUPS.push({ title: 'Form fields (/V)', items });
  }

  const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'];
  if (Array.isArray(kids)) {
    const items = [];
    kids.forEach((p, i) => {
      if (typeof p['/Rotate'] === 'number') {
        items.push(safeItem('rotate', `Page ${i+1} /Rotate`, () => p['/Rotate'], (nv)=>{ p['/Rotate'] = parseInt(nv,10) || 0; }));
      }
    });
    if (items.length) SAFE_GROUPS.push({ title: 'Page rotation', items });
  }

  if (Array.isArray(kids)) {
    const items = [];
    kids.forEach((p, i) => {
      const ann = p['/Annots'];
      if (!Array.isArray(ann)) return;
      ann.forEach((a, j) => {
        if (typeof a?.['/Contents'] === 'string') {
          items.push(safeItem('annot', `Page ${i+1} Annot #${j+1} /Contents`, () => a['/Contents'], (nv)=>{ a['/Contents'] = nv; }));
        }
      });
    });
    if (items.length) SAFE_GROUPS.push({ title: 'Annotation texts', items });
  }

  renderSafeGroups();
}

function walkFields(arr, visit, path='/AcroForm/Fields'){
  for (let i=0;i<arr.length;i++){
    const f = arr[i];
    if (f && typeof f === 'object') {
      visit(f, `${path}[${i}]`);
      if (Array.isArray(f['/Kids'])) walkFields(f['/Kids'], visit, `${path}[${i}]/Kids`);
    }
  }
}

function safeItem(kind, label, getter, setter){
  return {
    kind, label, getter, setter,
    preview(){ const v = getter(); return typeof v === 'string' ? v : JSON.stringify(v); }
  };
}

function renderSafeGroups(){
  const wrap = els.safeGroups; if (!wrap) return;
  const filter = (els.safeFilter?.value || '').toLowerCase();
  wrap.innerHTML = '';

  let total = 0;
  SAFE_GROUPS.forEach(group => {
    const visible = group.items.filter(it => !filter || it.label.toLowerCase().includes(filter) || it.preview().toLowerCase().includes(filter));
    if (!visible.length) return;
    total += visible.length;

    const box = document.createElement('div');
    box.className = 'text-group';
    box.innerHTML = `<h3>${group.title} — ${visible.length} item(s)</h3>`;

    visible.forEach(it => {
      const d = document.createElement('details');
      d.className = 'text-item';
      d.innerHTML = `
        <summary>
          <span class="ellipsis">${escapeHtml(it.label)}</span>
          <span class="meta"><span class="badge">${it.kind}</span></span>
        </summary>
        <div class="edit">
          <input class="sv" type="text" value="${escapeAttr(it.preview())}" />
          <div class="text-xy"></div>
          <div class="font-mini"></div>
          <div class="nudges"><button class="saveBtn">Apply</button></div>
        </div>
      `;
      d.querySelector('.saveBtn').onclick = async () => {
        const nv = d.querySelector('.sv').value;
        try {
          it.setter(nv);
          maybeRefreshJsonEditor(true);
          scheduleAssemble({ reason: 'global' });
        } catch (e) { console.error(e); toast('Failed to set value', true); }
      };
      box.appendChild(d);
    });

    wrap.appendChild(box);
  });

  if (els.safeStatus) els.safeStatus.textContent = `${total} editable item(s).`;
}

/* ---------------- Misc helpers ---------------- */
function escapeHtml(s){ return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function escapeAttr(s){ return escapeHtml(s); }
function fmtNum(v){ return Number.isFinite(v)? (Math.round(v*100)/100) : '—'; }
function numOrEmpty(v){ return Number.isFinite(v)? v : ''; }
function pdfEscapeLiteral(s){ return String(s).replace(/([()\\])/g,'\\$1'); }
function lastMatchBefore(regex, text, pos){ let m,last=null; const re=new RegExp(regex.source, regex.flags.includes('g')?regex.flags:regex.flags+'g'); while((m=re.exec(text)) && m.index<pos) last=m; return last; }
