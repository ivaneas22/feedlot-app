import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/feedlot-app/',
  build: {
    outDir: 'docs', // 👈 GitHub Pages busca la carpeta docs
  },
})
