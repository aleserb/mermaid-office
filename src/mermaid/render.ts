import DOMPurify from 'dompurify'
import mermaid from 'mermaid'
import type { DiagramTheme } from '../metadata/payload'

let renderSequence = 0

function initializeMermaid(theme: DiagramTheme) {
  mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'strict',
    theme,
    look: theme.startsWith('redux') || theme.startsWith('neo') ? 'neo' : 'classic',
    htmlLabels: false,
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
}

export async function renderMermaid(
  source: string,
  theme: DiagramTheme = 'default',
): Promise<string> {
  initializeMermaid(theme)

  if (!source.trim()) {
    throw new Error('Enter Mermaid diagram source.')
  }

  const id = `mermaid-office-${Date.now()}-${renderSequence++}`
  const { svg } = await mermaid.render(id, source)

  return DOMPurify.sanitize(svg, {
    USE_PROFILES: { svg: true, svgFilters: true },
    ADD_TAGS: ['style'],
  })
}
