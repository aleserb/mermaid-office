import DOMPurify from 'dompurify'
import mermaid from 'mermaid'

let initialized = false
let renderSequence = 0

function initializeMermaid() {
  if (initialized) {
    return
  }

  mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'strict',
    theme: 'default',
    htmlLabels: false,
    secure: [
      'secure',
      'securityLevel',
      'startOnLoad',
      'maxTextSize',
      'suppressErrorRendering',
      'maxEdges',
      'htmlLabels',
    ],
  })
  initialized = true
}

export async function renderMermaid(source: string): Promise<string> {
  initializeMermaid()

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
