// --- tiny CJS-ish polyfill for worker scope
self.global = self; // some deps check global
self.process = self.process || {
  env: {},
  nextTick: cb => Promise.resolve().then(cb)
};

import { PDFAssembler } from 'pdfassembler';

const ok = (id, payload) => self.postMessage({ id, ok: true, ...payload });
const fail = (id, err) =>
  self.postMessage({
    id, ok: false,
    error: (err && (err.stack || err.message || String(err))) || 'Unknown worker error'
  });

self.addEventListener('message', async (e) => {
  const { id, cmd } = e.data || {};
  try {
    if (cmd !== 'parse') return fail(id, 'unknown cmd');
    const { bytes } = e.data;
    const blob = new Blob([bytes], { type: 'application/pdf' });

    const asm = new PDFAssembler(blob);
    const pdfTree = await asm.getPDFStructure();
    let pageCount = 0;
    try { pageCount = await asm.countPages(); } catch {}

    ok(id, { pdfTree, pageCount });
  } catch (err) {
    fail(e.data?.id, err);
  }
});
