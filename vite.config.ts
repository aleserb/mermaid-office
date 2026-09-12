import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  base: '/mermaid-office/',
  plugins: [react()],
  build: {
    license: { fileName: 'third-party-licenses.txt' },
  },
})
