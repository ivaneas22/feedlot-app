import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/feedlot-app/', // 👈 importante: mismo nombre que el repo
})
