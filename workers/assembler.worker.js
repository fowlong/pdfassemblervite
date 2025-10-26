// --- tiny CJS-ish polyfill for worker scope (before any imports)
self.global = self;
self.process = self.process || {
  env: {},
  nextTick: (cb, ...args) => Promise.resolve().then(() => cb(...args))
};

// lazy-load pdfassembler so the worker can answer ping even if bundling is slow
let PDFAssemblerClass = null;
async function ensureLib() {
  if (!PDFAssemblerClass) {
    const mod = await import('pdfassembler');
    PDFAssemblerClass = mod.PDFAssembler;
  }
}

function ok(id, result, transfer) {
  // result should be a plain object
  if (transfer && transfer.length) {
    self.postMessage({ id, ok: true, result }, transfer);
  } else {
    self.postMessage({ id, ok: true, result });
  }
}
function fail(id, err) {
  const msg = (err && (err.stack || err.message)) || String(err || 'Unknown worker error');
  self.postMessage({ id, ok: false, error: msg });
}

// keep last assembled bytes so we can rehydrate context quickly
let lastBytes = null;

async function safeCountPages(asm) {
  try { return await asm.countPages(); } catch { return undefined; }
}

const handlers = {
  async ping() {
    return { pong: true };
  },

  // bytes:ArrayBuffer, options:{ compress:boolean, indent:boolean }
  async open({ bytes, options }) {
    await ensureLib();
    const u8 = new Uint8Array(bytes);
    lastBytes = u8; // cache
    const blob = new Blob([u8], { type: 'application/pdf' });

    const asm = new PDFAssemblerClass(blob);
    asm.compress = !!options?.compress;
    asm.indent   = options?.indent ? 2 : false;

    const tree = await asm.getPDFStructure();
    const pageCount = await safeCountPages(asm);
    return { treeJSON: JSON.stringify(tree), pageCount };
  },

  // treeJSON:string, options:{ compress, indent }
  async assemble({ treeJSON, options }) {
    await ensureLib();
    // Prefer building from previous bytes if we have them (better resource continuity),
    // but it also works with a fresh assembler.
    const asm = lastBytes
      ? new PDFAssemblerClass(new Blob([lastBytes], { type: 'application/pdf' }))
      : new PDFAssemblerClass();

    asm.compress = !!options?.compress;
    asm.indent   = options?.indent ? 2 : false;

    try { asm.pdfTree = JSON.parse(treeJSON); }
    catch { asm.pdfTree = treeJSON; }

    const out = await asm.assemblePdf('Uint8Array');
    lastBytes = out; // update cache with latest build
    // Transfer the buffer back
    return { __transfer__: [out.buffer], payload: { uint8: out.buffer } };
  },

  // bytes:ArrayBuffer
  async rehydrateFromBytes({ bytes }) {
    await ensureLib();
    const u8 = new Uint8Array(bytes);
    lastBytes = u8;
    const blob = new Blob([u8], { type: 'application/pdf' });
    const asm = new PDFAssemblerClass(blob);
    const tree = await asm.getPDFStructure();
    const pageCount = await safeCountPages(asm);
    return { treeJSON: JSON.stringify(tree), pageCount };
  }
};

self.addEventListener('message', async (e) => {
  const { id, type, payload } = e.data || {};
  if (!id || !type) return;

  const fn = handlers[type];
  if (!fn) return fail(id, `Unknown worker type "${type}"`);

  try {
    const res = await fn(payload || {});
    // Special casing transferables: if handler returned { __transfer__, payload }
    if (res && Array.isArray(res.__transfer__)) {
      const { __transfer__, payload } = res;
      return ok(id, payload, __transfer__);
    }
    ok(id, res);
  } catch (err) {
    fail(id, err);
  }
});
