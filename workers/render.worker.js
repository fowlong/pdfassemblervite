// render.worker.js
// Off-thread PDF.js render to OffscreenCanvas (requires COOP/COEP + crossOriginIsolated)

let pdfjsLib = null;
async function ensurePdfJs() {
  if (pdfjsLib) return pdfjsLib;
  pdfjsLib = await import('https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.mjs');
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.mjs';
  return pdfjsLib;
}

self.onmessage = async (e) => {
  const { id, type, payload } = e.data || {};
  const reply = (ok, result) => self.postMessage({ id, ok, ...(ok ? { result } : { error: String(result?.message || result) }) });

  try {
    if (type === 'render') {
      const { bytes, pageIndex=1, width=1000 } = payload || {};
      const canvas = e.ports?.[0] || null; // OffscreenCanvas passed via transfer (Chrome-style port), or payload.offscreen?
      let off = payload?.offscreen;
      if (!off && canvas && canvas instanceof OffscreenCanvas) off = canvas;
      if (!off && payload?.canvas) off = payload.canvas;
      if (!off) throw new Error('No OffscreenCanvas provided');

      await ensurePdfJs();
      const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
      const page = await pdf.getPage(Math.max(1, Math.min(pageIndex, pdf.numPages)));
      const viewport = page.getViewport({ scale: 1.0 });
      const scale = width / viewport.width;
      const vp = page.getViewport({ scale });
      off.width = vp.width; off.height = vp.height;

      const ctx = off.getContext('2d');
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      return reply(true, { width: off.width, height: off.height });
    }
    return reply(false, new Error('Unknown render command'));
  } catch (err) {
    return reply(false, err);
  }
};
