import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/feedlot-app/', // 👈 clave: esto le dice a GitHub dónde vive tu app
})
