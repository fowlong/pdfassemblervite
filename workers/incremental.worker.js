// workers/incremental.worker.js
// Incremental writer (v3): robust classic-xref and xref-stream appends.
// - Correct startxref offsets
// - Correctly bump /Size in trailer
// - Keeps internal trailer state in sync across multiple edits
// - Works whether the base file uses classic xref tables or xref streams.

let baseU8 = null;
let baseLen = 0;
let trailer = null;   // { prevXref, isXrefStream, size, rootRef, infoRef }
let pageMap = [];     // [{ pageObj:{n,g}, contents:[{n,g},...] }]

// --- helper: never transfer the live working buffer out of the worker
function snapshotBuffer(u8OrBuf) {
  if (u8OrBuf instanceof Uint8Array) return u8OrBuf.slice().buffer; // new buffer
  if (u8OrBuf && typeof u8OrBuf.byteLength === 'number' && typeof u8OrBuf.slice === 'function') {
    return u8OrBuf.slice(0); // ArrayBuffer clone
  }
  const view = new Uint8Array(u8OrBuf || 0);
  return view.slice().buffer;
}

self.onmessage = async (e) => {
  const { id, type, payload } = e.data || {};
  const ok   = (res, transfer=[]) => self.postMessage({ id, ok: true, result: res }, transfer);
  const fail = (err) => self.postMessage({ id, ok: false, error: (err && (err.stack || err.message)) || String(err) });

  try {
    if (type === 'open') {
      const { bytes, mapping } = payload || {};
      if (!(bytes instanceof ArrayBuffer)) throw new Error('open(): bytes must be ArrayBuffer');

      baseU8 = new Uint8Array(bytes);
      baseLen = baseU8.length;

      // Parse last trailer/xref
      const s = decodeLatin1(baseU8);
      trailer = parseLastTrailer(s); // detects classic vs xref-stream & reads Size/Root/Info/Prev

      // Optional mapping from main thread
      pageMap = [];
      if (mapping && Array.isArray(mapping.pages) && mapping.pages.length) {
        pageMap = mapping.pages.map(p => ({
          pageObj: p.pageObj || null,
          contents: Array.isArray(p.contents) ? p.contents.map(c => ({ n: c.n|0, g: c.g|0 })) : []
        }));
        if (!trailer.rootRef && mapping.root) trailer.rootRef = mapping.root;
        if (!trailer.infoRef && mapping.info) trailer.infoRef = mapping.info;
        if (!trailer.size && Number.isFinite(mapping.size)) trailer.size = mapping.size|0;
      } else {
        try { pageMap = tryBuildPagesFromClassicObjects(s, trailer.rootRef); } catch {}
      }

      if (!trailer.rootRef) throw new Error('open(): could not determine /Root reference');

      return ok({
        pageCount: pageMap.length,
        isXrefStream: trailer.isXrefStream === true
      });
    }

    if (type === 'getMapping') {
      return ok({ pages: pageMap });
    }

    if (type === 'applyEdits') {
      if (!baseU8 || !trailer) throw new Error('applyEdits(): call open() first.');
      const { edits } = payload || {};
      if (!Array.isArray(edits) || !edits.length) {
        const snap = snapshotBuffer(baseU8);
        return ok({ uint8: snap }, [snap]);
      }

      // Normalize edit targets to exact object ids
      const targets = [];
      for (const ed of edits) {
        if (ed && ed.obj && Number.isFinite(ed.obj.n)) {
          targets.push({ n: ed.obj.n|0, g: (ed.obj.g|0)||0, text: String(ed.streamText || '') });
        } else if (Number.isFinite(ed.pageIndex) && Number.isFinite(ed.streamIndex)) {
          const pg = pageMap[ed.pageIndex|0];
          const ref = pg?.contents?.[ed.streamIndex|0];
          if (!ref) throw new Error(`applyEdits(): invalid streamIndex ${ed.streamIndex} for page ${ed.pageIndex+1} (have ${pg?.contents?.length||0})`);
          targets.push({ n: ref.n|0, g: (ref.g|0)||0, text: String(ed.streamText || '') });
        } else {
          throw new Error('applyEdits(): edit must specify obj or (pageIndex, streamIndex)');
        }
      }

      // 1) append redefined objects (uncompressed)
      const parts = [];
      const offsets = new Map(); // objNum -> absolute offset
      let cursorAbs = baseLen;

      for (const t of targets) {
        const body = streamObj(t.n, t.g, t.text);
        offsets.set(t.n, cursorAbs);
        parts.push(body);
        cursorAbs += body.length;
      }

      // 2) write xref (classic table OR xref stream), set /Prev, and "startxref"
      let xrefStartAbs = cursorAbs; // we'll set precisely per-branch below
      let sizeNew = Math.max(trailer.size|0, maxObjNumInAppend(offsets) + 1);

      if (trailer.isXrefStream) {
        // ---- XRef stream append ----
        const maxObj = Math.max(maxObjNumInAppend(offsets), (trailer.size|0) - 1);
        const xrefObjNum = maxObj + 1;

        xrefStartAbs = cursorAbs; // xref stream begins here

        // Entries for changed objects + an entry for the xref stream itself
        const entries = [...offsets.entries()].map(([n, ofs]) => ({ n, t:1, ofs, gen:0 }));
        entries.push({ n: xrefObjNum, t:1, ofs: xrefStartAbs, gen:0 });

        sizeNew = Math.max(sizeNew, xrefObjNum + 1);

        const dict = {
          Size: sizeNew,
          Root: trailer.rootRef,
          Info: trailer.infoRef || null,
          Prev: trailer.prevXref|0
        };

        const xrefObjStr = buildXRefStreamObject(xrefObjNum, 0, entries, dict);
        parts.push(xrefObjStr);
        cursorAbs += xrefObjStr.length;

        const tail = `startxref\n${xrefStartAbs}\n%%EOF\n`;
        parts.push(tail);
        cursorAbs += tail.length;

        // keep trailer in sync for the next round
        trailer.prevXref = xrefStartAbs;
        trailer.size = sizeNew;
        trailer.isXrefStream = true;
      } else {
        // ---- Classic xref table append ----
        const xrefStartHere = cursorAbs; // beginning of "xref"
        const xrefTable = buildXrefTable(offsets);
        parts.push(xrefTable);
        cursorAbs += xrefTable.length;

        const trailerText = buildClassicTrailer({
          size: sizeNew,
          root: trailer.rootRef,
          info: trailer.infoRef,
          prev: trailer.prevXref|0
        }, xrefStartHere); // IMPORTANT: startxref must point to beginning of xref
        parts.push(trailerText);
        cursorAbs += trailerText.length;

        xrefStartAbs = xrefStartHere;

        // keep trailer in sync for the next round
        trailer.prevXref = xrefStartAbs;
        trailer.size = sizeNew;
        trailer.isXrefStream = false;
      }

      // 3) concat
      const appendStr = parts.join('');
      const out = new Uint8Array(baseLen + appendStr.length);
      out.set(baseU8, 0);
      encodeLatin1Into(appendStr, out, baseLen);

      // 4) cache
      baseU8 = out;
      baseLen = out.length;

      // 5) send a snapshot (NOT the live buffer)
      const snap = snapshotBuffer(out);
      return ok({ uint8: snap }, [snap]);
    }

    return fail(new Error('Unknown command: ' + type));
  } catch (err) {
    return fail(err);
  }
};

/* ---------------- helpers ---------------- */

function decodeLatin1(u8){ let s=''; for (let i=0;i<u8.length;i++) s+=String.fromCharCode(u8[i]); return s; }
function encodeLatin1Into(s, out, off){ for (let i=0;i<s.length;i++) out[off+i]=s.charCodeAt(i)&0xFF; }

function parseLastTrailer(s){
  const sx = s.lastIndexOf('startxref');
  if (sx < 0) throw new Error('startxref not found');
  const m = /startxref\s+(\d+)\s+%%EOF/gm.exec(s.slice(sx));
  if (!m) throw new Error('Could not parse startxref offset');
  const prevXref = parseInt(m[1], 10);

  // classic 'xref' vs xref stream
  const head = s.slice(prevXref, prevXref + 10);
  const isXrefStream = !/^xref[\r\n]/.test(head);

  if (!isXrefStream) {
    const trPos = s.indexOf('trailer', prevXref);
    if (trPos < 0) throw new Error('trailer not found after xref');
    const d0 = s.indexOf('<<', trPos);
    const d1 = findDictEnd(s, d0);
    const dict = s.slice(d0, d1);
    return {
      prevXref, isXrefStream: false,
      size: (+match1(/\/Size\s+(\d+)/, dict)) || 0,
      rootRef: refFrom(dict, '/Root'),
      infoRef: refFrom(dict, '/Info') || null
    };
  }

  // xref stream: prevXref points to "<n> <g> obj"
  const hdrM = /^(\d+)\s+(\d+)\s+obj\b/.exec(s.slice(prevXref, prevXref+40));
  if (!hdrM) throw new Error('xref-stream object header not found at startxref offset');
  const d0 = s.indexOf('<<', prevXref);
  const d1 = findDictEnd(s, d0);
  const dict = s.slice(d0, d1);

  return {
    prevXref, isXrefStream: true,
    size: (+match1(/\/Size\s+(\d+)/, dict)) || 0,
    rootRef: refFrom(dict, '/Root'),
    infoRef: refFrom(dict, '/Info') || null
  };
}
function findDictEnd(s, i0){ let d=0; for (let i=i0;i<s.length;i++){ if(s[i]==='<'&&s[i+1]==='<'){d++;i++;} else if(s[i]==='>'&&s[i+1]==='>'){d--;i++; if(d===0) return i+1;} } throw new Error('Unbalanced dict'); }
function match1(re, s){ const m=re.exec(s); return m?m[1]:null; }
function refFrom(dict, key){ const m=new RegExp(`${escapeRx(key)}\\s+(\\d+)\\s+(\\d+)\\s+R`).exec(dict); return m?{n:+m[1], g:+m[2]}:null; }
function escapeRx(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }

function tryBuildPagesFromClassicObjects(s, rootRef){
  if (!rootRef) return [];
  const map = indexClassicObjects(s);
  const root = map[rootRef.n]?.text || '';
  const pagesRef = refFrom(root, '/Pages');
  if (!pagesRef) return [];
  const out = []; walkPages(map, pagesRef, out);
  return out.map(p => ({ pageObj:{n:p.n,g:p.g}, contents: parsePageContentsRefs(map[p.n]?.text || '') }));
}
function indexClassicObjects(s){
  const out=Object.create(null); const re=/(\d+)\s+(\d+)\s+obj\b([\s\S]*?)endobj/g; let m;
  while((m=re.exec(s))){ out[+m[1]]={ g:+m[2], text:m[3] }; }
  return out;
}
function walkPages(objMap, ref, acc){
  const o = objMap[ref.n]; if(!o) return;
  const t = o.text;
  if (/\b\/Type\s*\/Page\b/.test(t)) { acc.push({n:ref.n, g:ref.g}); return; }
  if (!/\b\/Type\s*\/Pages\b/.test(t)) return;
  const kidsText = arrayTextOf(t, '/Kids');
  const kidRefs = refsInArray(kidsText);
  kidRefs.forEach(r => walkPages(objMap, r, acc));
}
function arrayTextOf(text, key){ const m=new RegExp(`${escapeRx(key)}\\s*\\[([\\s\\S]*?)\\]`).exec(text); return m?m[1]:''; }
function refsInArray(t){ const out=[]; const re=/(\d+)\s+(\d+)\s+R/g; let m; while((m=re.exec(t))) out.push({n:+m[1],g:+m[2]}); return out; }
function parsePageContentsRefs(pageDictText){
  const one=/\/Contents\s+(\d+)\s+(\d+)\s+R/.exec(pageDictText);
  if (one) return [{ n:+one[1], g:+one[2] }];
  const arr = arrayTextOf(pageDictText, '/Contents');
  if (!arr) return [];
  return refsInArray(arr);
}

function streamObj(n,g,plain){
  const len = plain.length;
  return `${n} ${g} obj\n<< /Length ${len} >>\nstream\n${plain}\nendstream\nendobj\n`;
}

function buildXrefTable(offsets){
  const nums = [...offsets.keys()].sort((a,b)=>a-b);
  const groups = groupContiguous(nums);
  let out = 'xref\n';
  for (const g of groups) {
    out += `${g[0]} ${g.length}\n`;
    for (const n of g) out += pad10(offsets.get(n)) + ' 00000 n \n';
  }
  return out;
}
function buildClassicTrailer({ size, root, info, prev }, xrefStart){
  const parts = [`/Size ${size||0}`, `/Root ${root.n} ${root.g} R`, `/Prev ${prev|0}`];
  if (info) parts.push(`/Info ${info.n} ${info.g} R`);
  const dict = `trailer\n<< ${parts.join(' ')} >>\n`;
  return `${dict}startxref\n${xrefStart}\n%%EOF\n`;
}

function buildXRefStreamObject(xrefN, xrefG, entries, dictExtras){
  // entries: [{n, t:1, ofs, gen}]
  const nums = entries.map(e=>e.n).sort((a,b)=>a-b);
  const groups = groupContiguous(nums);

  // W = [1, 8, 2] -> type (1 byte), offset (8 bytes), gen (2 bytes)
  const W = [1,8,2];
  const eByN = new Map(entries.map(e => [e.n, e]));

  const bodyParts = [];
  for (const g of groups) {
    for (let i=g[0]; i<g[0]+g.length; i++) {
      const e = eByN.get(i);
      bodyParts.push(binU(e.t|0, W[0]));
      bodyParts.push(binU(e.ofs|0, W[1]));
      bodyParts.push(binU(e.gen|0, W[2]));
    }
  }
  const body = joinUint8(bodyParts);
  const indexArr = groups.flatMap(g => [g[0], g.length]);

  const dictLines = [
    '/Type /XRef',
    `/W [${W.join(' ')}]`,
    `/Index [${indexArr.join(' ')}]`,
    `/Size ${dictExtras.Size||0}`,
    `/Root ${dictExtras.Root?.n||0} ${dictExtras.Root?.g||0} R`,
    ...(dictExtras.Info ? [`/Info ${dictExtras.Info.n} ${dictExtras.Info.g} R`] : []),
    `/Prev ${dictExtras.Prev|0}`,
    `/Length ${body.length}`
  ];
  const header = `${xrefN} ${xrefG} obj\n<< ${dictLines.join(' ')} >>\nstream\n`;
  const footer = `\nendstream\nendobj\n`;
  return header + latinFromU8(body) + footer;
}

function groupContiguous(nums){
  const out=[]; let cur=[];
  for (const n of nums) {
    if (!cur.length || n === cur[cur.length-1] + 1) cur.push(n);
    else { out.push(cur); cur=[n]; }
  }
  if (cur.length) out.push(cur);
  return out;
}
function pad10(n){ const s=String(n); return '0'.repeat(10 - s.length) + s; }
function binU(num, width){
  let n = Math.max(0, num >>> 0);
  const out = new Uint8Array(width);
  for (let i=width-1;i>=0;i--) { out[i] = n & 0xFF; n >>>= 8; }
  return out;
}
function joinUint8(arrs){
  const total = arrs.reduce((a,u)=>a+u.length,0);
  const out = new Uint8Array(total); let off=0;
  for (const u of arrs) { out.set(u, off); off += u.length; }
  return out;
}
function latinFromU8(u8){ let s=''; for (let i=0;i<u8.length;i++) s+=String.fromCharCode(u8[i]); return s; }
function maxObjNumInAppend(offsets){ return Math.max(...[...offsets.keys()], 0); }
