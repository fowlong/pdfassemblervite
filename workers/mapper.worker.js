// workers/mapper.worker.js
// Robust-enough byte mapper for incremental updates.
// Adds scanAllPages() so we can build a complete page->contents map even when the
// tree representation lacks object IDs.

const td = new TextDecoder('latin1');

function send(id, ok, result, error) {
  postMessage({ id, ok, result, error });
}

self.onmessage = async (e) => {
  const { id, type, payload } = e.data || {};
  try {
    if (type === 'contentsForPages') {
      const bytes = new Uint8Array(payload.bytes);
      const res = contentsForPages(bytes, payload.pages || []);
      send(id, true, res, null);
    } else if (type === 'scanAllPages') {
      const bytes = new Uint8Array(payload.bytes);
      const res = scanAllPages(bytes);
      send(id, true, res, null);
    } else {
      send(id, false, null, `unknown mapper call: ${type}`);
    }
  } catch (err) {
    send(id, false, null, err?.message || String(err));
  }
};

// Build mapping for a specific set of page object ids
function contentsForPages(bytes, pageIds) {
  const want = new Set(pageIds.map(p => `${p.n} ${p.g}`));
  const text = td.decode(bytes);

  const objRe = /(^|\n|\r)(\d+)\s+(\d+)\s+obj\b/g;
  const byKey = Object.create(null);

  while (true) {
    const m = objRe.exec(text);
    if (!m) break;
    const n = parseInt(m[2],10), g = parseInt(m[3],10);
    const off = m.index + (m[1] ? m[1].length : 0);
    const endIdx = text.indexOf('endobj', objRe.lastIndex);
    if (endIdx < 0) break;

    const key = `${n} ${g}`;
    if (!want.has(key)) { objRe.lastIndex = endIdx + 6; continue; }

    const slice = text.slice(off, endIdx + 6);
    if (!/\/Type\s*\/Page\b/.test(slice)) { objRe.lastIndex = endIdx + 6; continue; }

    byKey[key] = parseContents(slice);
    objRe.lastIndex = endIdx + 6;
  }

  return byKey;
}

// Scan ALL page objects in file order and extract their /Contents refs
function scanAllPages(bytes) {
  const text = td.decode(bytes);
  const objRe = /(^|\n|\r)(\d+)\s+(\d+)\s+obj\b/g;

  const pages = [];
  while (true) {
    const m = objRe.exec(text);
    if (!m) break;
    const n = parseInt(m[2],10), g = parseInt(m[3],10);
    const off = m.index + (m[1] ? m[1].length : 0);
    const endIdx = text.indexOf('endobj', objRe.lastIndex);
    if (endIdx < 0) break;

    const slice = text.slice(off, endIdx + 6);
    if (/\/Type\s*\/Page\b/.test(slice)) {
      pages.push({
        pageObj: { n, g },
        contents: parseContents(slice)
      });
    }
    objRe.lastIndex = endIdx + 6;
  }

  // Best-effort root/info discovery (optional)
  const root = findRef(text, /\/Root\s+(\d+)\s+(\d+)\s+R/);
  const info = findRef(text, /\/Info\s+(\d+)\s+(\d+)\s+R/);

  return { pages, root, info };
}

function parseContents(dictText) {
  const out = [];
  let m;

  // Array form
  const arrayRe = /\/Contents\s*\[\s*([^\]]+?)\s*\]/;
  m = arrayRe.exec(dictText);
  if (m) {
    const payload = m[1];
    const refRe = /(\d+)\s+(\d+)\s+R/g;
    while ((m = refRe.exec(payload))) {
      out.push({ n: parseInt(m[1],10), g: parseInt(m[2],10) });
    }
    return out;
  }

  // Single ref form
  const singleRe = /\/Contents\s+(\d+)\s+(\d+)\s+R/;
  m = singleRe.exec(dictText);
  if (m) {
    out.push({ n: parseInt(m[1],10), g: parseInt(m[2],10) });
  }
  return out;
}

function findRef(text, re){
  const m = re.exec(text);
  if (!m) return null;
  return { n: parseInt(m[1],10), g: parseInt(m[2],10) };
}
