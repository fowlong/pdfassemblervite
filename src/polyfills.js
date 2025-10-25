// minimal Node-ish globals for browser builds (for CJS deps like promise-queue)
if (typeof window !== 'undefined') {
  window.global = window.global || window;
  window.process = window.process || {
    env: {},
    nextTick: (cb, ...args) => Promise.resolve().then(() => cb(...args))
  };
}
