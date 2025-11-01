// src/workers/vector.worker.js
// Dedicated vector scan (paths/tables) off the main thread.
// Input: { type:'scanPaths', payload:{ pages:[{pageIndex, streams:[string]}], options:{ tablesMode:boolean } } }
// Output: { type:'scanPaths:done', results:[{pageIndex, items:[ ... ]}] }

self.onmessage = (e) => {
  const { type, payload } = e.data || {};
  if (type === 'scanPaths') {
    const { pages } = payload || {};
    const results = (pages || []).map(pg => scanPage(pg));
    self.postMessage({ type: 'scanPaths:done', results });
  }
};a

function scanPage(pg) {
  const out = { pageIndex: pg.pageIndex, items: [] };
  (pg.streams || []).forEach((src, streamIndex) => {
    if (typeof src !== 'string') return;
    const items = scanVectors(src).map(it => ({ ...it, streamIndex }));
    out.items.push(...items);
  });
  return out;
}

function scanVectors(src) {
  const items = [];

  // collect all cm
  const cms = [];
  const cmRe = /(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+cm/g;
  let m;
  while ((m = cmRe.exec(src))) {
    cms.push({ index: m.index, text: m[0], a:+m[1], b:+m[2], c:+m[3], d:+m[4], e:+m[5], f:+m[6] });
  }
  function lastCmBefore(pos) {
    let last = null;
    for (let i=0;i<cms.length;i++) { if (cms[i].index < pos) last = cms[i]; else break; }
    return last;
  }

  // rectangles
  const reRect = /(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+re/g;
  while ((m = reRect.exec(src))) {
    const start = m.index, end = reRect.lastIndex;
    const near = lastCmBefore(start);
    items.push({
      kind: 'rect', op: 're',
      start, end,
      hasCM: !!near, cmIndex: near?.index ?? null, cmText: near?.text ?? null,
      a: near?.a ?? 1, b: near?.b ?? 0, c: near?.c ?? 0, d: near?.d ?? 1, e: near?.e ?? 0, f: near?.f ?? 0,
      rect: { x:+m[1], y:+m[2], w:+m[3], h:+m[4] }
    });
  }

  // path paint ops
  const paintRe = /\b(S|s|f\*?|B\*?|b\*?|n)\b/g;
  while ((m = paintRe.exec(src))) {
    const start = m.index, end = paintRe.lastIndex;
    const near = lastCmBefore(start);
    items.push({
      kind: 'path', op: m[1],
      start, end,
      hasCM: !!near, cmIndex: near?.index ?? null, cmText: near?.text ?? null,
      a: near?.a ?? 1, b: near?.b ?? 0, c: near?.c ?? 0, d: near?.d ?? 1, e: near?.e ?? 0, f: near?.f ?? 0
    });
  }

  return items;
}
