import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  optimizeDeps: {
    include: ['jspdf', 'jspdf-autotable'],
    esbuildOptions: {
      plugins: []
    }
  },
  build: {
    commonjsOptions: {
      include: [/jspdf-autotable/, /node_modules/],
      transformMixedEsModules: true
    }
  },
  server: {
    port: 5173,
    proxy: {
      '/api': {
        target: 'http://localhost:5000',
        changeOrigin: true,
      }
    }
  }
})

