// scanner.worker.js
// Tokenize/scan Text items (Tj/TJ) and XObject draws per page.

function isWs(ch){ return ch===' '||ch==='\n'||ch==='\r'||ch==='\t'||ch==='\f'||ch==='\v'; }
function readNumber(src,i){ const m = /^-?\d+(?:\.\d+)?/.exec(src.slice(i)); if (!m) return null; return { num: parseFloat(m[0]), end: i+m[0].length }; }
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
function lastMatchBefore(regex, text, pos){ let m,last=null; const re=new RegExp(regex.source, regex.flags.includes('g')?regex.flags:regex.flags+'g'); while((m=re.exec(text)) && m.index<pos) last=m; return last; }
function pdfUnescapeLiteral(s){ return String(s).replace(/\\\\/g,'\u0000').replace(/\\\)/g,')').replace(/\\\(/g,'(').replace(/\u0000/g,'\\'); }

// Tj
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
// TJ
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

function findCmNear(src, pos, win=600){
  const start = Math.max(0, pos - win);
  const slice = src.slice(start, pos);
  const re = /(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+cm/g;
  let m, last=null;
  while ((m = re.exec(slice))) last = m;
  if (!last) return null;
  const [text,a,b,c,d,e,f] = last;
  return { index: start + last.index, text, a:Number(a), b:Number(b), c:Number(c), d:Number(d), e:Number(e), f:Number(f) };
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

/* ---------------- Commands ---------------- */

async function scanText({ treeJSON, pageIndex, xTol=4, yTol=2 }) {
  const tree = JSON.parse(treeJSON);
  const kids = tree?.['/Root']?.['/Pages']?.['/Kids'];
  const page = Array.isArray(kids) ? kids[pageIndex] : undefined;
  const groupsForPage = [];

  if (!page) return { pageIndex, items: [] };
  const contents = page?.['/Contents']; if(!contents) return { pageIndex, items: [] };
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
        pageIndex, streamIndex:sIndex,
        start:t.tokStart, end:t.tokEnd,
        text:t.text, posType, x, y, posIndex: posMatch ? posMatch.index : null,
        fontName: tf?tf[1]:null, fontSize: tf?Number(tf[2]):null,
        segments: t.segs
      };
    });

    // coalesce into groups
    let run = null;
    function flushRun(){
      if(!run) return;
      const first = run[0], last = run[run.length-1];
      groupsForPage.push({
        id: `p${pageIndex}s${first.streamIndex}o${first.start}`,
        pageIndex, streamIndex:first.streamIndex,
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
  return { pageIndex, items: groupsForPage };
}

async function scanXObjects({ treeJSON, pageIndex }) {
  const tree = JSON.parse(treeJSON);
  const kids = tree?.['/Root']?.['/Pages']?.['/Kids'];
  const page = Array.isArray(kids) ? kids[pageIndex] : undefined;
  const groups = [];
  if (!page) return { pageIndex, items: [] };

  const contents = page?.['/Contents']; if(!contents) return { pageIndex, items: [] };
  const streams = Array.isArray(contents) ? contents : [contents];

  streams.forEach((obj, sIndex) => {
    if (!obj || typeof obj !== 'object' || typeof obj.stream !== 'string') return;
    const src = obj.stream;
    const re = /\/([^\s]+)\s+Do/g;
    let m;
    while ((m = re.exec(src))) {
      const name = m[1];
      const doStart = m.index;
      const doEnd = re.lastIndex;
      const cm = findCmNear(src, doStart, 600);
      const kind = resolveXObjectKind(page, name) || 'Unknown';
      groups.push({
        pageIndex, streamIndex:sIndex,
        name, kind, doStart, doEnd,
        hasCM: !!cm,
        cmIndex: cm ? cm.index : null,
        cmText:  cm ? cm.text  : null,
        a: cm ? cm.a : 1, b: cm ? cm.b : 0, c: cm ? cm.c : 0, d: cm ? cm.d : 1, e: cm ? cm.e : 0, f: cm ? cm.f : 0
      });
    }
  });

  return { pageIndex, items: groups };
}

self.onmessage = async (e) => {
  const { id, type, payload } = e.data || {};
  const reply = (ok, result) => self.postMessage({ id, ok, ...(ok ? { result } : { error: String(result?.message || result) }) });

  try {
    if (type === 'scanText')  return reply(true, await scanText(payload));
    if (type === 'scanXObjects') return reply(true, await scanXObjects(payload));
    // future: scanPaths
    return reply(false, new Error('Unknown scanner command: ' + type));
  } catch (err) {
    return reply(false, err);
  }
};
