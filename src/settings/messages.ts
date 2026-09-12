import { isDiagramSettings, type DiagramKind, type DiagramSettings } from '../metadata/diagramSettings'
import { DIAGRAM_THEMES, type DiagramTheme } from '../metadata/payload'

export interface SettingsSnapshot {
  theme: DiagramTheme
  settings: DiagramSettings
  diagramKind: DiagramKind
}

export type SettingsMessage =
  | { type: 'ready' | 'cancel'; session: string }
  | ({ type: 'init'; session: string } & SettingsSnapshot)
  | { type: 'apply'; session: string; theme: DiagramTheme; settings: DiagramSettings }

export function parseSettingsMessage(message: string): SettingsMessage {
  const value: unknown = JSON.parse(message)
  if (!value || typeof value !== 'object') throw new Error('Invalid settings message.')
  const data = value as Record<string, unknown>
  if (typeof data.session !== 'string' || !data.session) throw new Error('Missing settings session.')
  if (data.type === 'ready' || data.type === 'cancel') {
    return { type: data.type, session: data.session }
  }
  const theme = DIAGRAM_THEMES.find(theme => theme === data.theme)
  if (!theme || !isDiagramSettings(data.settings)) throw new Error('Invalid diagram settings.')
  if (data.type === 'apply') {
    return { type: 'apply', session: data.session, theme, settings: data.settings }
  }
  const kinds: DiagramKind[] = ['flowchart', 'sequence', 'class', 'er', 'state', 'other']
  const diagramKind = kinds.find(kind => kind === data.diagramKind)
  if (data.type !== 'init' || !diagramKind) throw new Error('Invalid settings message type.')
  return { type: 'init', session: data.session, theme, settings: data.settings, diagramKind }
}
