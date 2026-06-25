import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import path from 'path'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
  server: {
    proxy: {
      // n8n is behind nginx on the public domain — proxy to domain, not raw IP
      '/proxy/n8n': { target: 'https://n8n.tikonacapital.com', rewrite: (p) => p.replace(/^\/proxy\/n8n/, ''), changeOrigin: true, secure: true },
      '/proxy/ppt': { target: 'http://72.61.226.16:8501',          rewrite: (p) => p.replace(/^\/proxy\/ppt/, ''), changeOrigin: true },
      '/proxy/fm':  { 
        target: 'http://72.61.226.16:8500',       
        rewrite: (p) => p.replace(/^\/proxy\/fm/,  ''), 
        changeOrigin: true,
        timeout: 900000,      // 15 minutes timeout
        proxyTimeout: 900000 // 15 minutes proxy timeout
      },
      '/proxy/goindia': {
        target: 'https://goindia-mcp.fly.dev',
        rewrite: (p) => p.replace(/^\/proxy\/goindia/, ''),
        changeOrigin: true,
        secure: true
      },
    },
  },
  build: {
    rollupOptions: {
      output: {
        manualChunks: {
          'vendor-react': ['react', 'react-dom', 'react-router-dom'],
          'vendor-query': ['@tanstack/react-query', '@tanstack/react-table'],
          'vendor-ui': [
            '@radix-ui/react-dialog',
            '@radix-ui/react-dropdown-menu',
            '@radix-ui/react-select',
            '@radix-ui/react-tabs',
            '@radix-ui/react-tooltip',
            '@radix-ui/react-alert-dialog',
          ],
          'vendor-supabase': ['@supabase/supabase-js'],
        },
      },
    },
  },
})
