// vite.config.js
import { defineConfig } from 'vite';

export default defineConfig({
  server: { open: '/index.html' },

  // Force Vite to prebundle CommonJS deps and provide globals
  optimizeDeps: {
    include: ['pdfassembler'],
    esbuildOptions: {
      define: {
        global: 'window',
        'process.env': '{}',
      },
    },
  },

  build: {
    commonjsOptions: {
      include: [/node_modules/],
      transformMixedEsModules: true,
    },
    rollupOptions: {
      output: { interop: 'auto' },
    },
  },

  // Prefer browser-friendly fields if the lib exposes them
  resolve: {
    mainFields: ['browser', 'module', 'jsnext:main', 'jsnext', 'main'],
  },
});
