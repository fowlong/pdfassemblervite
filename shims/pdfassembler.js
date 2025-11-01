// src/shims/pdfassembler.js
// Normalize the various ways "pdfassembler" might export its constructor,
// so both main thread and workers can import a single consistent value.

import * as ns from 'pdfassembler';

/** @type {any} */
let Ctor = null;

// 1) ESM named export: { PDFAssembler }
if (ns && typeof ns.PDFAssembler === 'function') {
  Ctor = ns.PDFAssembler;
}

// 2) ESM/CJS default export is the class itself
if (!Ctor && ns && typeof ns.default === 'function') {
  Ctor = ns.default;
}

// 3) Default export is an object containing the class: { default: { PDFAssembler } }
if (
  !Ctor &&
  ns &&
  ns.default &&
  typeof ns.default.PDFAssembler === 'function'
) {
  Ctor = ns.default.PDFAssembler;
}

// 4) CJS namespace itself is the constructor
if (!Ctor && typeof ns === 'function') {
  Ctor = ns;
}

if (!Ctor) {
  throw new Error(
    '[shim/pdfassembler] Could not locate PDFAssembler constructor in "pdfassembler" exports.'
  );
}

export { Ctor as PDFAssembler };
export default Ctor;
