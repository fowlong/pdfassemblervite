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
      els.jsonEditor.value = stringifyPdfTree(state.assembler.pdfTree);
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

  // Safe panel
  els.scanSafeBtn?.addEventListener('click', () => { if (!state.pdfTree) return toast('Load a PDF first.', true); scanSafeItems(); });
  els.safeFilter?.addEventListener('input', () => renderSafeGroups());
}

async function loadPdf(file) {
  resetPreview();
  toast('Loading…');
  try {
    state.assembler = new PDFAssembler(file);
    state.assembler.indent = els.indentToggle.checked ? 2 : false;
    state.assembler.compress = !!els.compressToggle.checked;

    const pdfTree = await state.assembler.getPDFStructure();
    state.pdfTree = pdfTree;

    try {
      state.pageCount = await state.assembler.countPages();
      els.pageIndex.max = state.pageCount || 1;
    } catch {}

    els.jsonEditor.value = stringifyPdfTree(pdfTree);
    els.assembleBtn.disabled = false;
    els.refreshPreviewBtn.disabled = false;
    els.syncFromTreeBtn.disabled = false;
    els.removeRootBtn.disabled = false;
    els.canvasModeBtn.disabled = false;

    setFrameUrl(URL.createObjectURL(file));
    toast('PDF loaded.');

    scanTextItems();
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
  }, 2);
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

async function assembleAndPreview() {
  if (!state.assembler) return;
  toast('Assembling…');
  const uint8 = await state.assembler.assemblePdf('Uint8Array');
  state.lastUint8 = uint8;
  const blob = new Blob([uint8], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  setFrameUrl(url);
  els.downloadBtn.disabled = false;
  els.downloadBtn.onclick = () => downloadBlob(blob, 'edited.pdf');
  toast('Assembled.');
}

async function autoAssembleAndRescan() {
  if (state.assembling) return;
  try {
    state.assembling = true;
    await assembleAndPreview();
    scanTextItems();
    scanSafeItems();
  } catch (e) {
    console.error(e);
    toast('Assemble failed. Likely invalid structure.', true);
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

  let changed = 0;
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
          const before = s.stream;
          try { s.stream = before.replace(re, repl); if (s.stream !== before) changed++; }
          catch (inner) { console.error('Replace failed on a stream:', inner); }
        }
      }
    }
    els.jsonEditor.value = stringifyPdfTree(state.pdfTree);
    autoAssembleAndRescan();
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
  if (!state.lastUint8) await autoAssembleAndRescan();
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
function drawOverlay(){ const ctx=els.overlayCanvas.getContext('2d'); ctx.clearRect(0,0,els.overlayCanvas.width,els.overlayCanvas.height); ctx.font='16px sans-serif'; ctx.fillStyle='#ffffff'; ctx.strokeStyle='#66ccff'; overlayItems.forEach(it=>{ ctx.fillText(it.text,it.x,it.y); const m=ctx.measureText(it.text); const w=m.width+8,h=22; ctx.strokeRect(it.x-4,it.y-16,w,h); }); }
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
// { kind:'Tj'|'TJ', tokStart, tokEnd, text, segs:[{start,end,text}] }
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
function isDigit(ch){ return ch>='0'&&ch<='9'; }
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
      // octal \ddd
      const oct = /^\\([0-7]{1,3})/.exec(src.slice(j));
      if (oct){ out += String.fromCharCode(parseInt(oct[1],8)); j += 1 + oct[1].length; continue; }
      out += nxt; j+=2; continue;
    } else if (ch === '('){ depth++; out+='('; j++; }
    else if (ch === ')'){ depth--; if (depth===0){ j++; break; } out+=')'; j++; }
    else { out += ch; j++; }
  }
  return { text: out, end: j };
}
// Read [ ... ] TJ where ... contains strings and numbers
function scanTokensTJ(src){
  const out = [];
  for (let i=0;i<src.length;i++){
    const ch = src[i];
    if (ch !== '[') continue;
    let j = i+1;
    const segs = [];
    // parse array body
    while (j < src.length) {
      while (j<src.length && isWs(src[j])) j++;
      if (src[j] === '(') {
        const s = readPdfString(src, j);
        if (!s) break;
        segs.push({ start: j, end: s.end, text: s.text });
        j = s.end;
      } else if (src[j] === ']') {
        j++;
        // skip spaces then require TJ
        let k=j; while (k<src.length && isWs(src[k])) k++;
        if (src.slice(k, k+2) === 'TJ') {
          const tokStart = i;
          const tokEnd = k+2;
          const text = segs.map(s=>s.text).join('');
          out.push({ kind:'TJ', tokStart, tokEnd, text, segs });
        }
        i = j; // continue outer for
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
/*
GroupItem {
  id, pageIndex, streamIndex, objRef,
  start, end, text, posType, x, y, posIndex, fontName, fontSize,
  segments: [{start,end,text}]  // literal segments inside this token/group
}
*/

function scanTextItems(){
  TEXT_GROUPS = [];
  const kids = state.pdfTree?.['/Root']?.['/Pages']?.['/Kids'];
  if(!Array.isArray(kids)){ toast('No /Pages./Kids found.', true); renderTextGroups(); return; }

  kids.forEach((page, pIndex) => {
    const contents = page?.['/Contents']; if(!contents) return;
    const streams = Array.isArray(contents) ? contents : [contents];
    const groupsForPage = [];

    streams.forEach((obj, sIndex) => {
      if (!obj || typeof obj !== 'object' || typeof obj.stream !== 'string') return;
      const src = obj.stream;

      const toks = [...scanTokensTj(src), ...scanTokensTJ(src)];
      toks.sort((a,b)=>a.tokStart-b.tokStart);

      // gather positioning near each token
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

      // coalesce adjacent items on same row and near X
      const xTol = Number(els.xTol?.value || 4);
      const yTol = Number(els.yTol?.value || 2);
      let run = null;
      function flushRun(){
        if(!run) return;
        const first = run[0], last = run[run.length-1];
        groupsForPage.push({
          id: `p${pIndex}s${sIndex}o${first.start}`,
          pageIndex:pIndex, streamIndex:sIndex, objRef:first.objRef,
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
    TEXT_GROUPS.push({ pageIndex:pIndex, items: groupsForPage });
  });

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

// Replace coalesced group with a safe single token: [(text)] TJ
async function replaceGroupText(it, newText){
  const obj = getObjRef(it); if(!obj) return toast('Stream not found.', true);
  const s = obj.stream;
  const replacement = `[(${pdfEscapeLiteral(newText)})] TJ`;
  obj.stream = s.slice(0, it.start) + replacement + s.slice(it.end);
  it.text = newText;
  els.jsonEditor.value = stringifyPdfTree(state.pdfTree);
  await autoAssembleAndRescan();
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
  els.jsonEditor.value = stringifyPdfTree(state.pdfTree);
  await autoAssembleAndRescan();
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
  els.jsonEditor.value = stringifyPdfTree(state.pdfTree);
  await autoAssembleAndRescan();
}

// ---------------------- Fields & Metadata (beta) -------------------------

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
          els.jsonEditor.value = stringifyPdfTree(state.pdfTree);
          await autoAssembleAndRescan();
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
