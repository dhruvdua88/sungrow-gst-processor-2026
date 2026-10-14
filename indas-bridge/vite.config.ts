import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Served at <pages-root>/indas-bridge/. Built after the main GST app (root `npm run build:all`),
// so the root build's emptyOutDir does not wipe this folder.
export default defineConfig({
  base: '/sungrow-gst-processor-2026/indas-bridge/',
  plugins: [react()],
  build: {
    outDir: '../dist/indas-bridge',
    emptyOutDir: true,
    chunkSizeWarningLimit: 600,
  },
  worker: { format: 'es' },
})
