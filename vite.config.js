import { defineConfig } from 'vite';
export default defineConfig({ root: process.env.VROOT || '.', base: './', build: { outDir: 'dist', emptyOutDir: true, chunkSizeWarningLimit: 20000 } });
