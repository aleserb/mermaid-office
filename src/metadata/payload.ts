export const DIAGRAM_SCHEMA_VERSION = 1
export const MERMAID_RENDERER_VERSION = '11.17.2'
export const CONTENT_CONTROL_TAG_PREFIX = `mermaid-office:v${DIAGRAM_SCHEMA_VERSION}:`
export const DOCUMENT_SETTING_PREFIX = 'mermaid-office:diagram:'

export type DiagramFormat = 'svg' | 'png'
export type DiagramSize = 'small' | 'medium' | 'large' | 'page-width'
export type DiagramTheme =
  | 'default'
  | 'neutral'
  | 'dark'
  | 'forest'
  | 'neo'
  | 'neo-dark'
  | 'redux'
  | 'redux-dark'
  | 'redux-color'
  | 'redux-dark-color'

export const DIAGRAM_THEMES: DiagramTheme[] = [
  'default',
  'neutral',
  'dark',
  'forest',
  'neo',
  'neo-dark',
  'redux',
  'redux-dark',
  'redux-color',
  'redux-dark-color',
]
export const DIAGRAM_SIZES: DiagramSize[] = [
  'small',
  'medium',
  'large',
  'page-width',
]

export interface DiagramPayload {
  schemaVersion: typeof DIAGRAM_SCHEMA_VERSION
  id: string
  source: string
  theme: DiagramTheme
  size: DiagramSize
  format: DiagramFormat
  rendererVersion: string
}

export function createDiagramPayload(
  source: string,
  format: DiagramFormat,
  theme: DiagramTheme = 'default',
  size: DiagramSize = 'medium',
): DiagramPayload {
  if (!source.trim()) {
    throw new Error('Mermaid source cannot be empty.')
  }

  return {
    schemaVersion: DIAGRAM_SCHEMA_VERSION,
    id: crypto.randomUUID(),
    source,
    theme,
    size,
    format,
    rendererVersion: MERMAID_RENDERER_VERSION,
  }
}

export function getContentControlTag(id: string): string {
  return `${CONTENT_CONTROL_TAG_PREFIX}${id}`
}

export function getDiagramIdFromTag(tag: string): string | null {
  if (!tag.startsWith(CONTENT_CONTROL_TAG_PREFIX)) {
    return null
  }

  const id = tag.slice(CONTENT_CONTROL_TAG_PREFIX.length)
  return id || null
}

export function getDocumentSettingKey(id: string): string {
  return `${DOCUMENT_SETTING_PREFIX}${id}`
}

export function parseDiagramPayload(value: string): DiagramPayload {
  let candidate: unknown

  try {
    candidate = JSON.parse(value)
  } catch {
    throw new Error('Diagram metadata is not valid JSON.')
  }

  if (!isDiagramPayload(candidate)) {
    throw new Error('Diagram metadata has an unsupported or invalid structure.')
  }

  return {
    ...candidate,
    size: candidate.size ?? 'medium',
  }
}

function isDiagramPayload(value: unknown): value is DiagramPayload {
  if (!value || typeof value !== 'object') {
    return false
  }

  const payload = value as Record<string, unknown>
  return (
    payload.schemaVersion === DIAGRAM_SCHEMA_VERSION &&
    typeof payload.id === 'string' &&
    payload.id.length > 0 &&
    typeof payload.source === 'string' &&
    payload.source.trim().length > 0 &&
    DIAGRAM_THEMES.includes(payload.theme as DiagramTheme) &&
    (payload.size === undefined ||
      DIAGRAM_SIZES.includes(payload.size as DiagramSize)) &&
    ['svg', 'png'].includes(String(payload.format)) &&
    typeof payload.rendererVersion === 'string'
  )
}
