import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/feedlot-app/', // 👈 clave para que GitHub Pages la encuentre
})
