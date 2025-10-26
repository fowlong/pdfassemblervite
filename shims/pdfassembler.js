// src/shims/pdfassembler.js
// Import the package once and normalize the export shape.
// This runs through Vite's CommonJS transform (because of optimizeDeps.include).
import * as _mod from 'pdfassembler';

// Possible shapes: {PDFAssembler}, default, or plain constructor.
const PDFAssembler =
  (_mod && _mod.PDFAssembler) ||
  (_mod && _mod.default && (_mod.default.PDFAssembler || _mod.default)) ||
  _mod.default ||
  _mod;

export { PDFAssembler };
export default PDFAssembler;
