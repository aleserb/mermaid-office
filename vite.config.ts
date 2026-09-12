import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { execFileSync } from 'node:child_process'
import { resolve } from 'node:path'

function getBuildVersion() {
  if (process.env.VITE_BUILD_VERSION) {
    return process.env.VITE_BUILD_VERSION
  }

  try {
    return execFileSync('git', ['rev-parse', '--short', 'HEAD'], {
      encoding: 'utf8',
    }).trim()
  } catch (error) {
    console.warn('Unable to read the Git revision for the editor build label.', error)
    return 'development'
  }
}

export default defineConfig({
  base: '/mermaid-office/',
  define: {
    __BUILD_VERSION__: JSON.stringify(getBuildVersion()),
  },
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
