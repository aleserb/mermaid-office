import DOMPurify from 'dompurify'
import type { MermaidConfig } from 'mermaid'
import { DEFAULT_DIAGRAM_SETTINGS, type DiagramKind, type DiagramSettings } from '../metadata/diagramSettings'
import type { DiagramTheme } from '../metadata/payload'

let renderSequence = 0
let mermaidPromise: Promise<typeof import('mermaid')['default']> | undefined
let elkRegistered = false
let renderQueue: Promise<void> = Promise.resolve()

export function detectDiagramKind(source: string): DiagramKind {
  let remaining = source.trimStart()
  let frontmatterSeen = false
  while (remaining) {
    if (remaining.startsWith('%%{')) {
      const end = remaining.indexOf('}%%')
      if (end < 0) return 'other'
      remaining = remaining.slice(end + 3).trimStart()
    } else if (remaining.startsWith('%%')) {
      const end = remaining.indexOf('\n')
      if (end < 0) return 'other'
      remaining = remaining.slice(end + 1).trimStart()
    } else if (!frontmatterSeen && /^---[ \t]*\r?\n/.test(remaining)) {
      const opening = remaining.indexOf('\n') + 1
      const end = /^---[ \t]*(?:\r?\n|$)/m.exec(remaining.slice(opening))
      if (!end) return 'other'
      remaining = remaining.slice(opening + end.index + end[0].length).trimStart()
      frontmatterSeen = true
    } else {
      break
    }
  }

  const header = /^([A-Za-z][A-Za-z0-9-]*)(?=\s|;|$)/.exec(remaining)?.[1]
  switch (header) {
    case 'graph':
    case 'flowchart':
    case 'flowchart-elk':
      return 'flowchart'
    case 'sequenceDiagram':
      return 'sequence'
    case 'classDiagram':
    case 'classDiagram-v2':
      return 'class'
    case 'erDiagram':
      return 'er'
    case 'stateDiagram':
    case 'stateDiagram-v2':
      return 'state'
    default:
      return 'other'
  }
}

async function loadMermaid() {
  mermaidPromise ??= import('mermaid').then((module) => module.default)
  return mermaidPromise
}

function getConfig(theme: DiagramTheme, settings: DiagramSettings, kind: DiagramKind): MermaidConfig {
  const secure = [
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
  ]
  const config: MermaidConfig = {
    startOnLoad: false,
    securityLevel: 'strict',
    theme,
    look: settings.look === 'auto'
      ? (theme.startsWith('redux') || theme.startsWith('neo') ? 'neo' : 'classic')
      : settings.look,
    htmlLabels: false,
    suppressErrorRendering: true,
    secure,
  }
  if (config.look === 'handDrawn') {
    // Keep the sketch stable when an edit or quality change regenerates the PNG.
    config.handDrawnSeed = 1
  }

  if (settings.fontFamily !== 'theme') {
    config.fontFamily = settings.fontFamily
    config.themeVariables = { fontFamily: settings.fontFamily }
    secure.push('fontFamily', 'actorFontFamily', 'noteFontFamily', 'messageFontFamily')
  }
  if (settings.fontSize !== 'theme') {
    config.fontSize = settings.fontSize
    config.themeVariables = { ...config.themeVariables, fontSize: `${settings.fontSize}px` }
    secure.push('fontSize', 'actorFontSize', 'noteFontSize', 'messageFontSize')
  }

  // Mermaid sanitizes secure keys recursively. Lock only selected leaves, not
  // entire diagram config objects, so unrelated source configuration still works.
  if (kind === 'flowchart') {
    const flowchart: NonNullable<MermaidConfig['flowchart']> = {}
    if (settings.spacing !== 'default') {
      flowchart.nodeSpacing = settings.spacing === 'compact' ? 25 : 80
      flowchart.rankSpacing = settings.spacing === 'compact' ? 25 : 80
      secure.push('nodeSpacing', 'rankSpacing')
    }
    if (settings.curve !== 'default') {
      flowchart.curve = settings.curve
      secure.push('curve')
    }
    if (settings.layout !== 'default') {
      config.layout = settings.layout
      flowchart.defaultRenderer = 'dagre-wrapper'
      secure.push('layout', 'defaultRenderer')
    }
    if (Object.keys(flowchart).length) config.flowchart = flowchart
  }

  if (kind === 'sequence') {
    // Legacy diagrams use default settings, so default-valued options must
    // still allow their existing frontmatter and init directives to take effect.
    const sequence: NonNullable<MermaidConfig['sequence']> = {}
    if (settings.sequenceNumbers !== DEFAULT_DIAGRAM_SETTINGS.sequenceNumbers) {
      sequence.showSequenceNumbers = settings.sequenceNumbers
      secure.push('showSequenceNumbers')
    }
    if (settings.sequenceWrap !== DEFAULT_DIAGRAM_SETTINGS.sequenceWrap) {
      sequence.wrap = settings.sequenceWrap
      secure.push('wrap')
    }
    if (settings.sequenceMirrorActors !== DEFAULT_DIAGRAM_SETTINGS.sequenceMirrorActors) {
      sequence.mirrorActors = settings.sequenceMirrorActors
      secure.push('mirrorActors')
    }
    if (settings.fontFamily !== 'theme') {
      sequence.actorFontFamily = settings.fontFamily
      sequence.noteFontFamily = settings.fontFamily
      sequence.messageFontFamily = settings.fontFamily
    }
    if (settings.fontSize !== 'theme') {
      sequence.actorFontSize = settings.fontSize
      sequence.noteFontSize = settings.fontSize
      sequence.messageFontSize = settings.fontSize
    }
    if (settings.spacing !== 'default') {
      sequence.actorMargin = settings.spacing === 'compact' ? 25 : 80
      sequence.messageMargin = settings.spacing === 'compact' ? 20 : 55
      secure.push('actorMargin', 'messageMargin')
    }
    if (Object.keys(sequence).length) config.sequence = sequence
  }

  return config
}

export async function renderMermaid(
  source: string,
  theme: DiagramTheme = 'default',
  settings: DiagramSettings = DEFAULT_DIAGRAM_SETTINGS,
): Promise<string> {
  if (!source.trim()) {
    throw new Error('Enter Mermaid diagram source.')
  }

  const config = getConfig(theme, settings, detectDiagramKind(source))
  const result = renderQueue.then(async () => {
    const mermaid = await loadMermaid()
    if (config.layout === 'elk' && !elkRegistered) {
      const { default: layouts } = await import('@mermaid-js/layout-elk')
      mermaid.registerLayoutLoaders(layouts)
      elkRegistered = true
    }
    mermaid.initialize(config)
    const id = `mermaid-office-${Date.now()}-${renderSequence++}`
    const { svg } = await mermaid.render(id, source)

    return DOMPurify.sanitize(svg, {
      USE_PROFILES: { svg: true, svgFilters: true },
      ADD_TAGS: ['style'],
    })
  })
  // Keep later renders usable after a failure; return the original rejection
  // to the caller. Initialization must share the queue with Mermaid's render.
  renderQueue = result.then(() => undefined, () => undefined)
  return result
}
