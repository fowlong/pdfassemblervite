import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const candidates = [
  path.join(root, 'node_modules', 'pdfassembler', 'dist', 'index.js'),
  path.join(root, 'node_modules', 'pdfassembler', 'dist', 'pdfassembler.js')
];

for (const f of candidates) {
  if (fs.existsSync(f)) {
    const src = fs.readFileSync(f, 'utf8');
    const out = src.replace(/\/\/# sourceMappingURL=.*$/gm, '');
    if (out !== src) {
      fs.writeFileSync(f, out);
      console.log('Patched sourcemap line in', f);
    }
  }
}
