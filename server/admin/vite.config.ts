import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// 开发时把 /admin/api 与 /api 转发给本地 FastAPI，生产由 nginx 反代，
// 这样浏览器始终只看到一个 origin，不需要为管理后台开 CORS。
export default defineConfig({
  plugins: [react()],
  server: {
    port: 8080,
    proxy: {
      '/admin/api': { target: 'http://127.0.0.1:8000', changeOrigin: true },
      '/api': { target: 'http://127.0.0.1:8000', changeOrigin: true },
    },
  },
  build: { outDir: 'dist' },
})
