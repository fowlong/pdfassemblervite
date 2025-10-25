// No pdfassembler import needed; this worker just runs pure scans
// Add your heavy text/xobject/vector scan code here later if you want
const ok = (id, payload) => self.postMessage({ id, ok: true, ...payload });
const fail = (id, err) =>
  self.postMessage({
    id, ok: false,
    error: (err && (err.stack || err.message || String(err))) || 'Unknown worker error'
  });

self.addEventListener('message', async (e) => {
  const { id, cmd } = e.data || {};
  try {
    if (cmd !== 'ping') return fail(id, 'unknown cmd');
    ok(id, { pong: true });
  } catch (err) {
    fail(id, err);
  }
});
