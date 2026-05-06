import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    // На Windows иногда слушали только [::1], а браузер ходит на 127.0.0.1 — страница «не грузится»
    host: true,
    port: 5173,
    strictPort: false,
  },
})
