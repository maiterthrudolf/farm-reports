import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';

export default defineConfig({
  plugins: [react()],
  base: '/farm-reports/',
  server: {
    port: 3000,
  },
  resolve: {
    alias: {
      // Use the browser-compatible ExcelJS bundle (no Node.js built-ins)
      exceljs: path.resolve(__dirname, 'node_modules/exceljs/dist/exceljs.bare.min.js'),
    },
  },
});
