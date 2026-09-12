export type DiagramKind = 'flowchart' | 'sequence' | 'class' | 'er' | 'state' | 'other'
export type ImageQuality = 'auto' | 'standard' | 'high'

export interface DiagramSettings {
  fontFamily: 'theme' | 'system-ui' | 'Arial' | 'Times New Roman' | 'monospace'
  fontSize: 'theme' | 12 | 14 | 16 | 18 | 20 | 24
  look: 'auto' | 'classic' | 'neo' | 'handDrawn'
  spacing: 'default' | 'compact' | 'spacious'
  curve: 'default' | 'linear' | 'basis' | 'step'
  sequenceNumbers: boolean
  sequenceWrap: boolean
  sequenceMirrorActors: boolean
  layout: 'default' | 'dagre' | 'elk'
  imageQuality: ImageQuality
}

export const DEFAULT_DIAGRAM_SETTINGS: DiagramSettings = Object.freeze({
  fontFamily: 'theme',
  fontSize: 'theme',
  look: 'auto',
  spacing: 'default',
  curve: 'default',
  sequenceNumbers: false,
  sequenceWrap: false,
  sequenceMirrorActors: true,
  layout: 'default',
  imageQuality: 'auto',
})

export function getDiagramSettings(settings?: DiagramSettings): DiagramSettings {
  return settings ?? DEFAULT_DIAGRAM_SETTINGS
}

export function sameDiagramSettings(left?: DiagramSettings, right?: DiagramSettings): boolean {
  const a = getDiagramSettings(left)
  const b = getDiagramSettings(right)
  return (Object.keys(DEFAULT_DIAGRAM_SETTINGS) as (keyof DiagramSettings)[])
    .every(key => a[key] === b[key])
}

export function isDiagramSettings(value: unknown): value is DiagramSettings {
  if (!value || typeof value !== 'object') return false
  const settings = value as Record<string, unknown>
  return typeof settings.fontFamily === 'string'
    && ['theme', 'system-ui', 'Arial', 'Times New Roman', 'monospace'].includes(settings.fontFamily)
    && (settings.fontSize === 'theme' || (typeof settings.fontSize === 'number' && [12, 14, 16, 18, 20, 24].includes(settings.fontSize)))
    && typeof settings.look === 'string' && ['auto', 'classic', 'neo', 'handDrawn'].includes(settings.look)
    && typeof settings.spacing === 'string' && ['default', 'compact', 'spacious'].includes(settings.spacing)
    && typeof settings.curve === 'string' && ['default', 'linear', 'basis', 'step'].includes(settings.curve)
    && typeof settings.sequenceNumbers === 'boolean'
    && typeof settings.sequenceWrap === 'boolean'
    && typeof settings.sequenceMirrorActors === 'boolean'
    && typeof settings.layout === 'string' && ['default', 'dagre', 'elk'].includes(settings.layout)
    && typeof settings.imageQuality === 'string' && ['auto', 'standard', 'high'].includes(settings.imageQuality)
}
