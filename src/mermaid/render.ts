import DOMPurify from 'dompurify'
import type { DiagramTheme } from '../metadata/payload'

let renderSequence = 0
let mermaidPromise: Promise<typeof import('mermaid')['default']> | undefined

async function loadMermaid() {
  mermaidPromise ??= import('mermaid').then((module) => module.default)
  return mermaidPromise
}

async function initializeMermaid(theme: DiagramTheme) {
  const mermaid = await loadMermaid()
  mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'strict',
    theme,
    look: theme.startsWith('redux') || theme.startsWith('neo') ? 'neo' : 'classic',
    htmlLabels: false,
    suppressErrorRendering: true,
    secure: [
      'secure',
      'securityLevel',
      'startOnLoad',
      'maxTextSize',
      'suppressErrorRendering',
      'maxEdges',
      'htmlLabels',
      'theme',
      'themeVariables',
      'look',
    ],
  })
  return mermaid
}

export async function renderMermaid(
  source: string,
  theme: DiagramTheme = 'default',
): Promise<string> {
  if (!source.trim()) {
    throw new Error('Enter Mermaid diagram source.')
  }

  const mermaid = await initializeMermaid(theme)
  const id = `mermaid-office-${Date.now()}-${renderSequence++}`
  const { svg } = await mermaid.render(id, source)

  return DOMPurify.sanitize(svg, {
    USE_PROFILES: { svg: true, svgFilters: true },
    ADD_TAGS: ['style'],
  })
}
