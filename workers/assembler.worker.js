// --- tiny CJS-ish polyfill for worker scope
self.global = self;
self.process = self.process || {
  env: {},
  nextTick: cb => Promise.resolve().then(cb)
};

import { PDFAssembler } from 'pdfassembler';

const ok = (id, payload, transfers) => self.postMessage({ id, ok: true, ...payload }, transfers || []);
const fail = (id, err) =>
  self.postMessage({
    id, ok: false,
    error: (err && (err.stack || err.message || String(err))) || 'Unknown worker error'
  });

self.addEventListener('message', async (e) => {
  const { id, cmd } = e.data || {};
  try {
    if (cmd !== 'assemble') return fail(id, 'unknown cmd');
    const { bytes, pdfTree, options } = e.data;

    const blob = new Blob([bytes], { type: 'application/pdf' });
    const asm = new PDFAssembler(blob);
    asm.compress = !!options?.compress;
    asm.indent   = options?.indent ? 2 : false;
    asm.pdfTree  = pdfTree;

    const out = await asm.assemblePdf('Uint8Array');
    ok(id, { bytes: out.buffer }, [out.buffer]); // transfer the buffer
  } catch (err) {
    fail(e.data?.id, err);
  }
});
