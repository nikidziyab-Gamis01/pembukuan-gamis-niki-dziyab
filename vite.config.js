import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev
export default defineConfig({
  plugins: [react()],
  base: './' // Baris sakti ini memastikan semua aset 404 terurai dengan benar di internet
})
