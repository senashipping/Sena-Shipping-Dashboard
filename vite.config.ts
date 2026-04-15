import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    open: true
  },
  publicDir: 'public',
  build: {
    outDir: 'dist',
    assetsDir: 'assets',
    copyPublicDir: true,
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes('node_modules/handsontable') || id.includes('node_modules/@handsontable')) {
            return 'handsontable';
          }
          if (id.includes('node_modules/exceljs')) {
            return 'exceljs';
          }
          if (id.includes('node_modules/chevrotain')) {
            return 'chevrotain';
          }
        }
      }
    }
  }
})