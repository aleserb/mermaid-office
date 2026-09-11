import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { resolve } from 'node:path'

export default defineConfig({
  base: '/mermaid-office/',
  build: {
    rolldownOptions: {
      input: {
        commands: resolve(import.meta.dirname, 'commands.html'),
        editor: resolve(import.meta.dirname, 'editor.html'),
        taskpane: resolve(import.meta.dirname, 'index.html'),
      },
    },
  },
  plugins: [react()],
})
