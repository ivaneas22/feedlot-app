import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/feedlot-app/', // 👈 esta ruta es la correcta para GitHub Pages
})
