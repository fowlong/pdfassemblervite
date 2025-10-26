// --- tiny polyfill so pdfassembler’s CJS deps don’t choke in the browser
if (typeof window !== 'undefined') {
  window.global = window.global || window;
  window.process = window.process || {
    env: {},
    nextTick: (cb, ...args) => Promise.resolve().then(() => cb(...args))
  };
}

import { PDFAssembler } from 'pdfassembler';

// ---- pdf.js for canvas preview (kept separate from pdfassembler’s v2)
let pdfjsLib = null;
async function ensurePdfJs() {
  if (pdfjsLib) return pdfjsLib;
  pdfjsLib = await import('https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.mjs');
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.mjs';
  return pdfjsLib;
}

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

let state = {
  assembler: null,
  pdfTree: null,
  lastBlobUrl: null,
  lastUint8: null,
  pageCount: 0,
  assembling: false
};

// --- FAST EDIT MODE knobs
const ASSEMBLE_DEBOUNCE_MS = 450;     // assemble less often
const JSON_REFRESH_MS = 1500;         // throttle JSON editor updates
let assembleTimer = null;
let idleRehydrateTimer = null;
const dirtyPages = new Set();
let lastJsonRefresh = 0;

function isNativeViewerActive() {
  const btn = document.querySelector('.tabs button[data-tab="native"]');
  return btn && btn.classList.contains('active');
}

// enqueue assemble + selective rescan; defer heavy rehydrate
function scheduleAssemble({ pageIndex = null, reason = 'text' } = {}) {
  if (pageIndex != null) dirtyPages.add(pageIndex);
  if (assembleTimer) clearTimeout(assembleTimer);

  assembleTimer = setTimeout(async () => {
    assembleTimer = null;

    // 1) assemble bytes fast (compression OFF while editing)
    await assembleAndPreviewConditional();

    // 2) lightweight: use current assembler.pdfTree (no rehydrate yet)
    state.pdfTree = state.assembler.pdfTree;

    // 3) selective rescan (pages we touched)
    const pages = [...dirtyPages];
    dirtyPages.clear();
    if (pages.length) {
      scanTextItems(pages);
      scanXObjects(pages);
      scanPaths(pages);
    } else {
      scanTextItems();
      scanXObjects();
      scanPaths();
    }

    // Safe panel only needs refresh when metadata/fields changed
    if (reason !== 'text') scanSafeItems();

    // 4) kick a background “truth sync” (rehydrate) on idle
    if (idleRehydrateTimer) clearTimeout(idleRehydrateTimer);
    idleRehydrateTimer = setTimeout(async () => {
      await rehydrateTreeFromBytes(); // full parse from bytes
      if (pages.length) { scanTextItems(pages); scanXObjects(pages); scanPaths(pages); }
      maybeRefreshJsonEditor(true);
    }, 1200);
  }, ASSEMBLE_DEBOUNCE_MS);
}

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
  // Paths/SVG/Tables panel (NEW)
  scanPathsBtn: $('#scanPathsBtn'),
  pathFilter: $('#pathFilter'),
  pNudgeStep: $('#pNudgeStep'),
  pScaleStep: $('#pScaleStep'),
  tablesMode: $('#tablesMode'),
  pathStatus: $('#pathStatus'),
  pathGroups: $('#pathGroups'),
  // Safe panel
  scanSafeBtn: $('#scanSafeBtn'),
  safeFilter: $('#safeFilter'),
  safeStatus: $('#safeStatus'),
  safeGroups: $('#safeGroups'),
};

wireUi();
toast('Ready. Load a PDF to begin.');

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

  els.indentToggle.addEventListener('change', () => {
    if (state.assembler) state.assembler.indent = els.indentToggle.checked ? 2 : false;
  });
  els.compressToggle.addEventListener('change', () => {
    if (state.assembler) state.assembler.compress = !!els.compressToggle.checked;
  });

  els.assembleBtn.addEventListener('click', () => autoAssembleAndRescan());

  els.removeRootBtn.addEventListener('click', async () => {
    if (!state.assembler) return;
    try {
      state.assembler.removeRootEntries(['/Outlines', '/PageLabels']);
      maybeRefreshJsonEditor(true);
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

  // Paths panel (NEW)
  els.scanPathsBtn?.addEventListener('click', () => { if (!state.pdfTree) return toast('Load a PDF first.', true); scanPaths(); });
  els.pathFilter?.addEventListener('input', () => renderPathGroups());
  els.tablesMode?.addEventListener('change', () => renderPathGroups());

  // Safe panel
  els.scanSafeBtn?.addEventListener('click', () => { if (!state.pdfTree) return toast('Load a PDF first.', true); scanSafeItems(); });
  els.safeFilter?.addEventListener('input', () => renderSafeGroups());
}

async function loadPdf(file) {
  resetPreview();
  toast('Loading…');
  try {
    state.assembler = new PDFAssembler(file);

    // start in fast mode for editing
    els.compressToggle.checked = false;
    els.indentToggle.checked = false;
    state.assembler.compress = false;
    state.assembler.indent = false;

    const pdfTree = await state.assembler.getPDFStructure();
    state.pdfTree = pdfTree;

    try {
      state.pageCount = await state.assembler.countPages();
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

    scanTextItems();
    scanXObjects();
    scanPaths();       // <— NEW
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
      if (seen.has(value)) return { $ref: true }; // avoid cycles
      seen.add(value);
    }
    return value;
  }, 2);
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
  if (!state.assembler) return;
  try {
    const edited = JSON.parse(els.jsonEditor.value);
    state.assembler.pdfTree = edited;
    state.pdfTree = edited;
    if (andAssemble) await autoAssembleAndRescan();
  } catch (e) {
    console.error(e);
    toast('JSON parse failed. Fix errors and try again.', true);
  }
}

async function assembleAndPreviewConditional() {
  if (!state.assembler) return;
  const uint8 = await state.assembler.assemblePdf('Uint8Array');
  state.lastUint8 = uint8;

  if (isNativeViewerActive()) {
    const blob = new Blob([uint8], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    setFrameUrl(url);
    els.downloadBtn.disabled = false;
    els.downloadBtn.onclick = () => downloadBlob(blob, 'edited.pdf');
  }
  toast('Assembled.');
}

// --- reconstruct assembler from the bytes we just built, then parse structure
async function rehydrateTreeFromBytes() {
  if (!state.lastUint8) return;
  try {
    const blob = new Blob([state.lastUint8], { type: 'application/pdf' });
    const asm = new PDFAssembler(blob);
    asm.indent = els.indentToggle.checked ? 2 : false;
    asm.compress = !!els.compressToggle.checked;

    const fresh = await asm.getPDFStructure();
    state.assembler = asm;
    state.pdfTree = fresh;
  } catch (e) {
    console.warn('rehydrate from bytes failed:', e);
  }
}

// Toolbar button: force full reparse & full rescan
async function autoAssembleAndRescan() {
  if (state.assembling) return;
  try {
    state.assembling = true;
    await assembleAndPreviewConditional();
    await rehydrateTreeFromBytes();
    scanTextItems();
    scanXObjects();
    scanPaths();       // <— NEW
    scanSafeItems();
    maybeRefreshJsonEditor(true);
  } catch (e) {
    console.error(e);
    toast('Assemble/rescan failed. See console.', true);
  } finally {
    state.assembling = false;
  }
}

function setFrameUrl(url) {
  if (state.lastBlobUrl) URL.revokeObjectURL(state.lastBlobUrl);
  state.lastBlobUrl = url;
  els.pdfFrame.src = url;
}
function downloadBlob(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove(); }
function toast(msg, error=false) { els.status.textContent = msg; els.status.style.color = error ? '#ffb0b0' : '#cce7ff'; if (!error) console.log(msg); }

// ---------- Quick Action: regex replace ----------
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

// ---------- Canvas Mode (pdf.js + overlay) ----------
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
function clearOverlay() { const octx = els.overlayCanvas.getContext('2d'); octx.clearRect(0,0,els.overlayCanvas.width, els.overlayCanvas.height); overlayItems.length = 0; }

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

// ----------------------- Text items (Tj + TJ) ------------------------

// escape/unescape for PDF literal strings
function pdfEscapeLiteral(s){ return String(s).replace(/([()\\])/g,'\\$1'); }
function pdfUnescapeLiteral(s){ return String(s).replace(/\\\\/g,'\u0000').replace(/\\\)/g,')').replace(/\\\(/g,'(').replace(/\u0000/g,'\\'); }
// last match before index
function lastMatchBefore(regex, text, pos){ let m,last=null; const re=new RegExp(regex.source, regex.flags.includes('g')?regex.flags:regex.flags+'g'); while((m=re.exec(text)) && m.index<pos) last=m; return last; }

// token model for one text drawing op
function scanTokensTj(src) {
  const out = [];
  const re = /\((?:\\.|[^\\()])*\)\s*Tj/g;
  let m;
  while ((m = re.exec(src))) {
    const start = m.index;
    const end   = re.lastIndex;
    const mm = /^\(((?:\\.|[^\\()])*)\)\s*Tj$/.exec(m[0]);
    if (!mm) continue;
    out.push({
      kind: 'Tj',
      tokStart: start,
      tokEnd: end,
      text: pdfUnescapeLiteral(mm[1]),
      segs: [{ start, end, text: pdfUnescapeLiteral(mm[1]) }]
    });
  }
  return out;
}

function isWs(ch){ return ch===' '||ch==='\n'||ch==='\r'||ch==='\t'||ch==='\f'||ch==='\v'; }
function readNumber(src,i){
  const m = /^-?\d+(?:\.\d+)?/.exec(src.slice(i));
  if (!m) return null;
  return { num: parseFloat(m[0]), end: i+m[0].length };
}
function readPdfString(src,i){
  if (src[i] !== '(') return null;
  let j=i+1, depth=1, out='';
  while (j < src.length && depth>0) {
    const ch = src[j];
    if (ch === '\\') {
      const nxt = src[j+1];
      if (nxt === '(' || nxt === ')' || nxt === '\\') { out += nxt; j+=2; continue; }
      if (nxt === 'n'){ out+='\n'; j+=2; continue; }
      if (nxt === 'r'){ out+='\r'; j+=2; continue; }
      if (nxt === 't'){ out+='\t'; j+=2; continue; }
      if (nxt === 'b'||nxt==='f'){ j+=2; continue; }
      const oct = /^\\([0-7]{1,3})/.exec(src.slice(j));
      if (oct){ out += String.fromCharCode(parseInt(oct[1],8)); j += 1 + oct[1].length; continue; }
      out += nxt; j+=2; continue;
    } else if (ch === '('){ depth++; out+='('; j++; }
    else if (ch === ')'){ depth--; if (depth===0){ j++; break; } out+=')'; j++; }
    else { out += ch; j++; }
  }
  return { text: out, end: j };
}
// Read [ ... ] TJ allowing whitespace/newlines between ']' and 'TJ'
function scanTokensTJ(src){
  const out = [];
  for (let i=0;i<src.length;i++){
    if (src[i] !== '[') continue;
    let j = i+1, segs = [];
    while (j < src.length) {
      while (j<src.length && isWs(src[j])) j++;
      if (src[j] === '(') {
        const s = readPdfString(src, j);
        if (!s) break;
        segs.push({ start: j, end: s.end, text: s.text });
        j = s.end;
      } else if (src[j] === ']') {
        j++;
        let k=j; while (k<src.length && isWs(src[k])) k++;
        if (src.slice(k, k+2) === 'TJ') {
          out.push({ kind:'TJ', tokStart:i, tokEnd:k+2, text: segs.map(s=>s.text).join(''), segs });
        }
        i = j;
        break;
      } else {
        const num = readNumber(src, j);
        if (num) { j = num.end; } else { break; }
      }
    }
  }
  return out;
}

// live model
let TEXT_GROUPS = []; // [{pageIndex, items:[GroupItem]}]

function scanTextItems(onlyPages = null){
  const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'];
  const allPages = Array.isArray(kids) ? kids.map((_,i)=>i) : [];
  const pages = Array.isArray(onlyPages) && onlyPages.length ? onlyPages : allPages;

  const buckets = new Map();

  if (Array.isArray(TEXT_GROUPS) && TEXT_GROUPS.length && onlyPages) {
    for (const g of TEXT_GROUPS) buckets.set(g.pageIndex, { pageIndex:g.pageIndex, items:[...g.items] });
  }

  for (const pIndex of pages) {
    const page = kids?.[pIndex];
    const groupsForPage = [];
    const contents = page?.['/Contents']; if(!contents) { buckets.set(pIndex,{pageIndex:pIndex,items:[]}); continue; }
    const streams = Array.isArray(contents) ? contents : [contents];

    streams.forEach((obj, sIndex) => {
      if (!obj || typeof obj !== 'object' || typeof obj.stream !== 'string') return;
      const src = obj.stream;

      const toks = [...scanTokensTj(src), ...scanTokensTJ(src)];
      toks.sort((a,b)=>a.tokStart-b.tokStart);

      const items = toks.map(t => {
        const tm = lastMatchBefore(/(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+Tm/g, src, t.tokStart);
        const td = lastMatchBefore(/(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+Td/g, src, t.tokStart);
        const tf = lastMatchBefore(/\/([^\s]+)\s+(-?\d+(?:\.\d+)?)\s+Tf/g, src, t.tokStart);
        let posType=null,x=NaN,y=NaN,posMatch=null;
        if (tm && (!td || tm.index > td.index)) { posType='Tm'; x=Number(tm[5]); y=Number(tm[6]); posMatch=tm; }
        else if (td) { posType='Td'; x=Number(td[1]); y=Number(td[2]); posMatch=td; }
        return {
          pageIndex:pIndex, streamIndex:sIndex, objRef:obj,
          start:t.tokStart, end:t.tokEnd,
          text:t.text, posType, x, y, posIndex: posMatch ? posMatch.index : null,
          fontName: tf?tf[1]:null, fontSize: tf?Number(tf[2]):null,
          segments: t.segs
        };
      });

      const xTol = Number(els.xTol?.value || 4);
      const yTol = Number(els.yTol?.value || 2);
      let run = null;
      function flushRun(){
        if(!run) return;
        const first = run[0], last = run[run.length-1];
        groupsForPage.push({
          id: `p${pIndex}s${first.streamIndex}o${first.start}`,
          pageIndex:pIndex, streamIndex:first.streamIndex, objRef:first.objRef,
          start:first.start, end:last.end,
          text: run.map(r=>r.text).join(''),
          posType:first.posType, x:first.x, y:first.y, posIndex:first.posIndex,
          fontName:first.fontName, fontSize:first.fontSize,
          segments: run.flatMap(r=>r.segments)
        });
        run=null;
      }
      for (const it of items){
        if (!run){ run=[it]; continue; }
        const prev = run[run.length-1];
        const sameRow = Number.isFinite(it.y) && Number.isFinite(prev.y) ? Math.abs(it.y - prev.y) <= yTol : true;
        const closeX  = (Number.isFinite(it.x) && Number.isFinite(prev.x)) ? (it.x - prev.x) <= xTol : true;
        if (it.streamIndex === prev.streamIndex && sameRow && closeX) run.push(it);
        else { flushRun(); run=[it]; }
      }
      flushRun();
    });

    groupsForPage.sort((a,b)=> (b.y ?? 0) - (a.y ?? 0));
    buckets.set(pIndex, { pageIndex:pIndex, items:groupsForPage });
  }

  const maxPage = (kids?.length || 0) - 1;
  TEXT_GROUPS = [];
  for (let i=0;i<=maxPage;i++){
    TEXT_GROUPS.push(buckets.get(i) || { pageIndex:i, items:[] });
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
        ${it.fontSize? `<span class="badge">size:${fmtNum(it.fontSize)}</span>`:''}
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
        <label>Size: <input class="ti-fs" type="number" step="0.5" value="${it.fontSize ?? ''}"></label>
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

function getObjRef(it){
  if (it?.objRef && typeof it.objRef.stream === 'string') return it.objRef;
  const page = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids']?.[it.pageIndex];
  const contents = page?.['/Contents'];
  const streams = Array.isArray(contents) ? contents : [contents];
  const obj = streams[it.streamIndex];
  return (obj && typeof obj.stream === 'string') ? obj : null;
}

// Replace group with safe single token; re-find locally if needed
async function replaceGroupText(it, newText){
  const obj = getObjRef(it); if(!obj) return toast('Stream not found.', true);
  const s = obj.stream;
  const replacement = `[(${pdfEscapeLiteral(newText)})] TJ`;

  const slice = s.slice(it.start, it.end);
  if (slice.length && s.indexOf(slice, Math.max(0, it.start - 4)) >= 0) {
    obj.stream = s.slice(0, it.start) + replacement + s.slice(it.end);
  } else {
    const winStart = Math.max(0, it.start - 400);
    const winEnd   = Math.min(s.length, it.end + 400);
    const win = s.slice(winStart, winEnd);
    const local = (() => {
      const i = win.lastIndexOf('[', it.start - winStart);
      let j = -1;
      if (i >= 0) {
        const close = win.indexOf(']', i + 1);
        if (close > i) {
          let k = close + 1; while (k < win.length && isWs(win[k])) k++;
          if (win.slice(k, k + 2) === 'TJ') j = k + 2;
        }
      }
      if (i >= 0 && j > i) return { a: winStart + i, b: winStart + j };
      const k = win.lastIndexOf('(', it.start - winStart);
      const l = win.indexOf(') Tj', Math.max(k + 1, 0));
      if (k >= 0 && l > k) return { a: winStart + k, b: winStart + l + 3 };
      return null;
    })();
    if (!local) return toast('Could not locate token to replace.', true);
    obj.stream = s.slice(0, local.a) + replacement + s.slice(local.b);
  }

  it.text = newText;
  maybeRefreshJsonEditor(true);
  scheduleAssemble({ pageIndex: it.pageIndex, reason: 'text' });
}

async function nudgeItem(it, dx, dy){
  if (!it.posType || it.posIndex == null) return toast('No position operator near this text; cannot nudge.', true);
  const obj = getObjRef(it); if(!obj) return toast('Stream not found.', true);
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
  const obj = getObjRef(it); if(!obj) return toast('Stream not found.', true);
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

// ---------------------- XObjects / Images (position / size) -------------------------

/*
We scan for `/Name Do` draws. For each, we capture the nearest preceding
`a b c d e f cm` (within a local window). Edits rewrite that `cm` or inject one.
*/

let XOBJ_GROUPS = []; // [{pageIndex, items:[XItem]}]
/*
XItem {
  pageIndex, streamIndex, objRef, name, kind, // kind: 'Image' | 'Form' | 'Unknown'
  doStart, doEnd,
  hasCM, cmIndex, cmText, a,b,c,d,e,f
}
*/

function scanXObjects(onlyPages = null){
  XOBJ_GROUPS = [];
  const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'];
  const allPages = Array.isArray(kids) ? kids.map((_,i)=>i) : [];
  const pages = Array.isArray(onlyPages) && onlyPages.length ? onlyPages : allPages;

  const buckets = new Map();
  if (onlyPages && XOBJ_GROUPS.length) {
    for (const g of XOBJ_GROUPS) buckets.set(g.pageIndex, { pageIndex:g.pageIndex, items:[...g.items] });
  }

  for (const pIndex of pages) {
    const page = kids?.[pIndex];
    const groups = [];
    const contents = page?.['/Contents']; if(!contents) { buckets.set(pIndex,{pageIndex:pIndex,items:[]}); continue; }
    const streams = Array.isArray(contents) ? contents : [contents];

    streams.forEach((obj, sIndex) => {
      if (!obj || typeof obj !== 'object' || typeof obj.stream !== 'string') return;
      const src = obj.stream;

      // find /Name Do
      const re = /\/([^\s]+)\s+Do/g;
      let m;
      while ((m = re.exec(src))) {
        const name = m[1]; // like Im1
        const doStart = m.index;
        const doEnd = re.lastIndex;

        // nearest cm within a short window before Do
        const cm = findCmNear(src, doStart, 600);
        const kind = resolveXObjectKind(page, name) || 'Unknown';

        groups.push({
          pageIndex: pIndex, streamIndex: sIndex, objRef: obj,
          name, kind, doStart, doEnd,
          hasCM: !!cm,
          cmIndex: cm ? cm.index : null,
          cmText:  cm ? cm.text  : null,
          a: cm ? cm.a : 1, b: cm ? cm.b : 0, c: cm ? cm.c : 0, d: cm ? cm.d : 1, e: cm ? cm.e : 0, f: cm ? cm.f : 0
        });
      }
    });

    buckets.set(pIndex, { pageIndex:pIndex, items:groups });
  }

  const maxPage = (kids?.length || 0) - 1;
  XOBJ_GROUPS = [];
  for (let i=0;i<=maxPage;i++){
    XOBJ_GROUPS.push(buckets.get(i) || { pageIndex:i, items:[] });
  }

  renderXGroups();
}

function resolveXObjectKind(page, shortName){
  try{
    const res = page?.['/Resources']?.['/XObject'];
    if (!res || typeof res !== 'object') return null;
    const key = `/${shortName}`;
    const dict = res[key];
    const sub = dict?.['/Subtype'];
    if (sub === '/Image') return 'Image';
    if (sub === '/Form') return 'Form';
    return null;
  }catch{ return null; }
}

// find the last "a b c d e f cm" within ~win chars before pos
function findCmNear(src, pos, win=500){
  const start = Math.max(0, pos - win);
  const slice = src.slice(start, pos);
  const re = /(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+cm/g;
  let m, last=null;
  while ((m = re.exec(slice))) last = m;
  if (!last) return null;
  const [text,a,b,c,d,e,f] = last;
  return { index: start + last.index, text, a:Number(a), b:Number(b), c:Number(c), d:Number(d), e:Number(e), f:Number(f) };
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
    box.className = 'text-group'; // reuse styling
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
      ${axisAligned ? '' : `<div style="opacity:.8;font-size:12px;margin-top:4px">Non-axis-aligned (b/c ≠ 0). You can still edit all 6 cm values manually:</div>
        <div class="text-xy">
          <label>a: <input class="xo-a" type="number" step="0.01" value="${it.a}"></label>
          <label>b: <input class="xo-b" type="number" step="0.01" value="${it.b}"></label>
          <label>c: <input class="xo-c" type="number" step="0.01" value="${it.c}"></label>
          <label>d: <input class="xo-d" type="number" step="0.01" value="${it.d}"></label>
          <label>e: <input class="xo-e" type="number" step="0.5" value="${it.e}"></label>
          <label>f: <input class="xo-f" type="number" step="0.5" value="${it.f}"></label>
          <button class="xo-apply6">Apply 6</button>
        </div>`}
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
  if (q('.xo-apply6')) {
    q('.xo-apply6').onclick = async () => {
      const a = Number(q('.xo-a').value), b = Number(q('.xo-b').value),
            c = Number(q('.xo-c').value), d = Number(q('.xo-d').value),
            e = Number(q('.xo-e').value), f = Number(q('.xo-f').value);
      if (![a,b,c,d,e,f].every(Number.isFinite)) return toast('Enter valid numbers', true);
      await xApplyMatrix(it, a,b,c,d,e,f);
    };
  }
  q('.xo-x').onchange = async () => { const tx = Number(q('.xo-x').value); if(Number.isFinite(tx)) await xApplyMatrix(it, it.a,it.b,it.c,it.d,tx,it.f); };
  q('.xo-y').onchange = async () => { const ty = Number(q('.xo-y').value); if(Number.isFinite(ty)) await xApplyMatrix(it, it.a,it.b,it.c,it.d,it.e,ty); };

  return d;
}

function getStreamForX(it){
  if (it?.objRef && typeof it.objRef.stream === 'string') return it.objRef;
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
    // replace in-place
    obj.stream = s.slice(0, it.cmIndex) + token + s.slice(it.cmIndex + it.cmText.length);
  } else {
    // inject a cm just before the Do
    obj.stream = s.slice(0, it.doStart) + token + ' ' + s.slice(it.doStart);
  }

  it.hasCM = true; it.cmText = token; it.a=a; it.b=b; it.c=c; it.d=d; it.e=e; it.f=f;
  maybeRefreshJsonEditor(true);
  scheduleAssemble({ pageIndex: it.pageIndex, reason: 'text' });
}
async function xNudge(it, dx, dy){ await xApplyMatrix(it, it.a, it.b, it.c, it.d, it.e + dx, it.f + dy); }
async function xScale(it, sx, sy){ await xApplyMatrix(it, it.a * sx, it.b, it.c, it.d * sy, it.e, it.f); }

// ---------------------- SVG / PATHS / TABLES (NEW) -------------------------

/*
We detect:
- Rectangles:    "x y w h re"
- Generic paths: from a 'm' (moveTo) until the next paint op S|s|f|f*|B|B*|b|b*|n
We expose / edit the nearest preceding "a b c d e f cm", or inject one.
*/

let PATH_GROUPS = [];  // [{ pageIndex, items:[PItem] }]
/*
PItem {
  kind: 'rect' | 'path',
  pageIndex, streamIndex, objRef,
  // shared matrix
  hasCM, cmIndex, cmText, a,b,c,d,e,f,
  // rect-specific
  reStart, reEnd, x, y, w, h, paint: 'S'|'s'|'f'|'f*'|'B'|'B*'|'b'|'b*'|'n'|null,
  // path-specific
  pStart, pEnd, paintOp, segCount, preview
}
*/

function scanPaths(onlyPages=null){
  PATH_GROUPS = [];
  const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'];
  const allPages = Array.isArray(kids) ? kids.map((_,i)=>i) : [];
  const pages = Array.isArray(onlyPages) && onlyPages.length ? onlyPages : allPages;

  const buckets = new Map();

  for (const pIndex of pages) {
    const page = kids?.[pIndex];
    const items = [];
    const contents = page?.['/Contents']; if(!contents) { buckets.set(pIndex, {pageIndex:pIndex, items:[]}); continue; }
    const streams = Array.isArray(contents) ? contents : [contents];

    streams.forEach((obj, sIndex) => {
      if (!obj || typeof obj !== 'object' || typeof obj.stream !== 'string') return;
      const src = obj.stream;

      // Rectangles: x y w h re
      const reRe = /(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+re/g;
      let rm;
      while ((rm = reRe.exec(src))) {
        const reStart = rm.index, reEnd = reRe.lastIndex;
        const cm = findCmNear(src, reStart, 600);
        const paint = detectPaintOpAfter(src, reEnd);
        items.push({
          kind:'rect', pageIndex:pIndex, streamIndex:sIndex, objRef:obj,
          reStart, reEnd,
          x:Number(rm[1]), y:Number(rm[2]), w:Number(rm[3]), h:Number(rm[4]),
          paint,
          hasCM: !!cm, cmIndex: cm?cm.index:null, cmText: cm?cm.text:null,
          a: cm?cm.a:1, b: cm?cm.b:0, c: cm?cm.c:0, d: cm?cm.d:1, e: cm?cm.e:0, f: cm?cm.f:0
        });
      }

      // Generic paths: from 'm' until paint op
      const mRe = /(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+m/g;
      let mm;
      while ((mm = mRe.exec(src))) {
        const mIdx = mm.index;
        const paint = findNextPaint(src, mRe.lastIndex);
        if (!paint) continue;
        const pStart = mIdx, pEnd = paint.end;
        const cm = findCmNear(src, pStart, 600);
        const segmentInfo = summarizePath(src.slice(pStart, pEnd));
        items.push({
          kind:'path', pageIndex:pIndex, streamIndex:sIndex, objRef:obj,
          pStart, pEnd, paintOp: paint.op, segCount: segmentInfo.count, preview: segmentInfo.preview,
          hasCM: !!cm, cmIndex: cm?cm.index:null, cmText: cm?cm.text:null,
          a: cm?cm.a:1, b: cm?cm.b:0, c: cm?cm.c:0, d: cm?cm.d:1, e: cm?cm.e:0, f: cm?cm.f:0
        });
        mRe.lastIndex = pEnd; // skip to after this painted path
      }
    });

    buckets.set(pIndex, { pageIndex:pIndex, items });
  }

  // flatten to array ordered by pages
  const maxPage = (state.pdfTree?.['/Root']?.['/Pages']?.['/Kids']?.length || 0) - 1;
  PATH_GROUPS = [];
  for (let i=0;i<=maxPage;i++){
    PATH_GROUPS.push(buckets.get(i) || { pageIndex:i, items:[] });
  }

  renderPathGroups();
}

function detectPaintOpAfter(src, from){
  const m = /(\bf\*?\b|\bB\*?\b|\bb\*?\b|\bS\b|\bs\b|\bn\b)/.exec(src.slice(from));
  return m ? m[1] : null;
}
function findNextPaint(src, from){
  const re = /(\bf\*?\b|\bB\*?\b|\bb\*?\b|\bS\b|\bs\b|\bn\b)/g;
  re.lastIndex = from;
  const m = re.exec(src);
  if (!m) return null;
  return { op: m[1], start: from + m.index, end: from + re.lastIndex };
}
function summarizePath(segText){
  const counts = {
    m: (segText.match(/\bm\b/g)||[]).length,
    l: (segText.match(/\bl\b/g)||[]).length,
    c: (segText.match(/\bc\b/g)||[]).length,
    h: (segText.match(/\bh\b/g)||[]).length,
    re:(segText.match(/\bre\b/g)||[]).length
  };
  const count = counts.m + counts.l + counts.c + counts.h + counts.re;
  const prev = segText.replace(/\s+/g,' ').trim().slice(0,120);
  return { count, preview: prev };
}

function renderPathGroups(){
  const wrap = els.pathGroups; if (!wrap) return;
  const filter = (els.pathFilter?.value || '').toLowerCase();
  const step   = Number(els.pNudgeStep?.value || 1);
  const sStep  = Number(els.pScaleStep?.value || 0.05);
  const tables = !!els.tablesMode?.checked;

  wrap.innerHTML = '';
  let total = 0;

  PATH_GROUPS.forEach(group=>{
    let items = group.items;
    if (filter) {
      items = items.filter(it => {
        const label = it.kind === 'rect'
          ? `rect ${it.paint||''}`
          : `path ${it.paintOp||''} segs:${it.segCount}`;
        return label.toLowerCase().includes(filter);
      });
    }
    total += items.length;

    const box = document.createElement('div');
    box.className = 'text-group';
    box.innerHTML = `<h3>Page ${group.pageIndex+1} — ${items.length} vector item(s)</h3>`;

    // Tables mode: group rects by similar Y
    if (tables) {
      const rects = items.filter(it=>it.kind==='rect');
      const groups = groupRectsAsTable(rects, 2 /*yTol*/, 2 /*xTol*/);
      groups.forEach((g,gi)=>{
        const row = document.createElement('details');
        row.className = 'text-item';
        row.innerHTML = `
          <summary><span class="ellipsis">Row ${gi+1}: ${g.items.length} cell(s)</span>
          <span class="meta"><span class="badge">bulk</span></span></summary>
          <div class="edit">
            <div class="nudges">
              <button class="nL">←</button>
              <button class="nU">↑</button>
              <button class="nD">↓</button>
              <button class="nR">→</button>
            </div>
          </div>`;
        const q = sel => row.querySelector(sel);
        q('.nL').onclick = async () => { await bulkRectNudge(g.items, -step, 0, group.pageIndex); };
        q('.nR').onclick = async () => { await bulkRectNudge(g.items, +step, 0, group.pageIndex); };
        q('.nU').onclick = async () => { await bulkRectNudge(g.items, 0, +step, group.pageIndex); };
        q('.nD').onclick = async () => { await bulkRectNudge(g.items, 0, -step, group.pageIndex); };
        box.appendChild(row);
      });
    }

    items.forEach(it => box.appendChild(renderPathItem(it, step, sStep)));
    wrap.appendChild(box);
  });

  if (els.pathStatus) els.pathStatus.textContent = `${total} vector item(s).`;
}

function renderPathItem(it, step, sStep){
  const d = document.createElement('details');
  d.className = 'text-item';

  const axisAligned = Math.abs(it.b) < 1e-6 && Math.abs(it.c) < 1e-6;
  const meta = it.kind==='rect'
    ? `rect ${it.paint||''} x:${fmtNum(it.x)} y:${fmtNum(it.y)} w:${fmtNum(it.w)} h:${fmtNum(it.h)}`
    : `path ${it.paintOp||''} segs:${it.segCount}`;

  d.innerHTML = `
    <summary>
      <span class="ellipsis">${escapeHtml(meta)}</span>
      <span class="meta">
        <span class="badge">p${it.pageIndex+1}</span>
        <span class="badge">${it.hasCM ? 'cm' : 'no cm'}</span>
        <span class="badge">tx:${fmtNum(it.e)} ty:${fmtNum(it.f)}</span>
        ${axisAligned ? `<span class="badge">sx:${fmtNum(it.a)} sy:${fmtNum(it.d)}</span>` : `<span class="badge">matrix</span>`}
      </span>
    </summary>
    <div class="edit">
      ${it.kind==='rect' ? `
      <div class="text-xy">
        <label>X: <input class="pr-x" type="number" step="0.5" value="${numOrEmpty(it.x)}"></label>
        <label>Y: <input class="pr-y" type="number" step="0.5" value="${numOrEmpty(it.y)}"></label>
        <label>W: <input class="pr-w" type="number" step="0.5" value="${numOrEmpty(it.w)}"></label>
        <label>H: <input class="pr-h" type="number" step="0.5" value="${numOrEmpty(it.h)}"></label>
        <button class="pr-apply">Apply rect</button>
      </div>` : `
      <div class="text-xy">
        <div class="mono small" style="opacity:.8;max-width:100%;overflow-x:auto">${escapeHtml((it.preview||'').trim())}</div>
      </div>`}
      <div class="text-xy">
        <label>tx: <input class="pm-x" type="number" step="0.5" value="${numOrEmpty(it.e)}"></label>
        <label>ty: <input class="pm-y" type="number" step="0.5" value="${numOrEmpty(it.f)}"></label>
        <span style="opacity:.7;font-size:12px">(cm)</span>
      </div>
      <div class="font-mini">
        <label>sx: <input class="pm-sx" type="number" step="0.01" value="${it.a}"></label>
        <label>sy: <input class="pm-sy" type="number" step="0.01" value="${it.d}"></label>
        <button class="pm-apply" title="Apply matrix">Apply</button>
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
      ${!axisAligned ? `<div class="text-xy">
        <label>a: <input class="pm-a" type="number" step="0.01" value="${it.a}"></label>
        <label>b: <input class="pm-b" type="number" step="0.01" value="${it.b}"></label>
        <label>c: <input class="pm-c" type="number" step="0.01" value="${it.c}"></label>
        <label>d: <input class="pm-d" type="number" step="0.01" value="${it.d}"></label>
        <label>e: <input class="pm-e" type="number" step="0.5" value="${it.e}"></label>
        <label>f: <input class="pm-f" type="number" step="0.5" value="${it.f}"></label>
        <button class="pm-apply6">Apply 6</button>
      </div>`:''}
    </div>
  `;

  const q = sel => d.querySelector(sel);

  // rect direct apply
  if (it.kind==='rect') {
    q('.pr-apply').onclick = async () => {
      const x = Number(q('.pr-x').value), y = Number(q('.pr-y').value),
            w = Number(q('.pr-w').value), h = Number(q('.pr-h').value);
      if (![x,y,w,h].every(Number.isFinite)) return toast('Enter valid rect', true);
      await pApplyRect(it, x,y,w,h);
    };
  }

  // matrix nudges/scales
  q('.nL').onclick = async () => { await pNudge(it, -step, 0); };
  q('.nR').onclick = async () => { await pNudge(it, +step, 0); };
  q('.nU').onclick = async () => { await pNudge(it, 0, +step); };
  q('.nD').onclick = async () => { await pNudge(it, 0, -step); };
  q('.sUp').onclick = async () => { await pScale(it, 1+sStep, 1+sStep); };
  q('.sDown').onclick = async () => { await pScale(it, 1/(1+sStep), 1/(1+sStep)); };
  q('.reset').onclick = async () => { await pApplyMatrix(it, 1,0,0,1,0,0); };

  q('.pm-apply').onclick = async () => {
    const sx = Number(q('.pm-sx').value); const sy = Number(q('.pm-sy').value);
    const tx = Number(q('.pm-x').value);  const ty = Number(q('.pm-y').value);
    if (![sx,sy,tx,ty].every(Number.isFinite)) return toast('Enter valid numbers', true);
    await pApplyMatrix(it, sx, 0, 0, sy, tx, ty);
  };
  if (q('.pm-apply6')) {
    q('.pm-apply6').onclick = async () => {
      const a = Number(q('.pm-a').value), b = Number(q('.pm-b').value),
            c = Number(q('.pm-c').value), d = Number(q('.pm-d').value),
            e = Number(q('.pm-e').value), f = Number(q('.pm-f').value);
      if (![a,b,c,d,e,f].every(Number.isFinite)) return toast('Enter valid numbers', true);
      await pApplyMatrix(it, a,b,c,d,e,f);
    };
  }
  q('.pm-x').onchange = async () => { const tx = Number(q('.pm-x').value); if(Number.isFinite(tx)) await pApplyMatrix(it, it.a,it.b,it.c,it.d,tx,it.f); };
  q('.pm-y').onchange = async () => { const ty = Number(q('.pm-y').value); if(Number.isFinite(ty)) await pApplyMatrix(it, it.a,it.b,it.c,it.d,it.e,ty); };

  return d;
}

// ---------- low-level mutators for paths ----------

function getStreamForPath(it){
  if (it?.objRef && typeof it.objRef.stream === 'string') return it.objRef;
  const page = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids']?.[it.pageIndex];
  const contents = page?.['/Contents']; const streams = Array.isArray(contents) ? contents : [contents];
  const obj = streams[it.streamIndex];
  return (obj && typeof obj.stream === 'string') ? obj : null;
}
function fm(n){ return (Math.round(n*10000)/10000).toString(); }

async function pApplyMatrix(it, a,b,c,d,e,f){
  const obj = getStreamForPath(it); if(!obj) return toast('Stream not found.', true);
  const s = obj.stream;
  const token = `${fm(a)} ${fm(b)} ${fm(c)} ${fm(d)} ${fm(e)} ${fm(f)} cm`;

  const insertAt = it.kind==='rect' ? it.reStart : it.pStart;

  if (it.hasCM && it.cmIndex != null && typeof it.cmText === 'string') {
    obj.stream = s.slice(0, it.cmIndex) + token + s.slice(it.cmIndex + it.cmText.length);
  } else {
    obj.stream = s.slice(0, insertAt) + token + ' ' + s.slice(insertAt);
  }

  it.hasCM = true; it.cmText = token; it.a=a; it.b=b; it.c=c; it.d=d; it.e=e; it.f=f;
  maybeRefreshJsonEditor(true);
  scheduleAssemble({ pageIndex: it.pageIndex, reason: 'text' });
}
async function pNudge(it, dx, dy){ await pApplyMatrix(it, it.a, it.b, it.c, it.d, it.e + dx, it.f + dy); }
async function pScale(it, sx, sy){ await pApplyMatrix(it, it.a * sx, it.b, it.c, it.d * sy, it.e, it.f); }

async function pApplyRect(it, x,y,w,h){
  const obj = getStreamForPath(it); if(!obj) return toast('Stream not found.', true);
  const s = obj.stream;
  const repl = `${fm(x)} ${fm(y)} ${fm(w)} ${fm(h)} re`;
  obj.stream = s.slice(0, it.reStart) + repl + s.slice(it.reEnd);
  it.x=x; it.y=y; it.w=w; it.h=h; it.reEnd = it.reStart + repl.length;
  maybeRefreshJsonEditor(true);
  scheduleAssemble({ pageIndex: it.pageIndex, reason: 'text' });
}

// ---------- very-light tables grouping (rows by Y) ----------

function groupRectsAsTable(rects, yTol=2, xTol=2){
  rects = rects.slice().sort((a,b)=> (b.y - a.y)); // top-down
  const rows = [];
  rects.forEach(r=>{
    let row = rows.find(R=>Math.abs(R.refY - r.y) <= yTol);
    if (!row){ row = { refY: r.y, items: [] }; rows.push(row); }
    row.items.push(r);
  });
  rows.forEach(row => row.items.sort((a,b)=> a.x - b.x));
  return rows;
}

async function bulkRectNudge(items, dx, dy, pageIndex){
  for (const it of items){
    if (it.hasCM) await pNudge(it, dx, dy);
    else await pApplyRect(it, it.x + dx, it.y + dy, it.w, it.h);
  }
  scheduleAssemble({ pageIndex, reason:'text' });
}

// ---------------------- Fields & Metadata (beta) -------------------------

let SAFE_GROUPS = []; // array of {title, items:[{kind, label, get(), set(newVal), preview()}]}

function scanSafeItems(){
  SAFE_GROUPS = [];

  // Document Info dictionary
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

  // AcroForm text fields (/V as string)
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

  // Per-page rotation
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

  // Annotation texts (/Contents)
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

// -------------------- helpers --------------------
function escapeHtml(s){ return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function escapeAttr(s){ return escapeHtml(s); }
function fmtNum(v){ return Number.isFinite(v)? (Math.round(v*100)/100) : '—'; }
function numOrEmpty(v){ return Number.isFinite(v)? v : ''; }
function resetPreview(){ if(state.lastBlobUrl) URL.revokeObjectURL(state.lastBlobUrl); state.lastBlobUrl=null; state.lastUint8=null; els.pdfFrame.removeAttribute('src'); }
