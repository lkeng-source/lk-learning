import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// GitHub Pages 路徑為 /lk-learning/
export default defineConfig({
  plugins: [react()],
  base: '/lk-learning/',
  server: {
    port: 5173,
    host: true
  }
})
