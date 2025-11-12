import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  base: '/feedlot-app/',   // <- ruta del repo para GH Pages
  plugins: [react()],
})
