import DOMPurify from 'dompurify'
import mermaid from 'mermaid'
import type { DiagramTheme } from '../metadata/payload'

let renderSequence = 0

function initializeMermaid(theme: DiagramTheme) {
  const isRedux = theme === 'redux'
  mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'strict',
    theme: isRedux ? 'base' : theme,
    ...(isRedux && {
      themeVariables: {
        background: '#ffffff',
        primaryColor: '#764abc',
        primaryBorderColor: '#4f326f',
        primaryTextColor: '#ffffff',
        secondaryColor: '#ede9f5',
        secondaryBorderColor: '#764abc',
        secondaryTextColor: '#2f2140',
        tertiaryColor: '#f7f4fb',
        tertiaryBorderColor: '#b39ddb',
        tertiaryTextColor: '#2f2140',
        lineColor: '#4f326f',
        textColor: '#2f2140',
      },
    }),
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
