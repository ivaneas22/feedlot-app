import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  base: '/feedlot-app/', // 👈 esta línea es CLAVE
  plugins: [react()],
})
