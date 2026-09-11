import {
  DIAGRAM_SIZES,
  DIAGRAM_THEMES,
  type DiagramSize,
  type DiagramTheme,
} from '../metadata/payload'

export type ParentToDialogMessage = {
  type: 'initialize'
  source: string
  theme: DiagramTheme
  size: DiagramSize
}

export type DialogToParentMessage =
  | { type: 'ready' }
  | { type: 'save'; source: string; theme: DiagramTheme; size: DiagramSize }
  | { type: 'cancel' }

export function parseParentMessage(value: string): ParentToDialogMessage {
  const message: unknown = JSON.parse(value)
  if (
    !message ||
    typeof message !== 'object' ||
    (message as Record<string, unknown>).type !== 'initialize' ||
    typeof (message as Record<string, unknown>).source !== 'string' ||
    !DIAGRAM_THEMES.includes(
      (message as Record<string, unknown>).theme as DiagramTheme,
    ) ||
    !DIAGRAM_SIZES.includes((message as Record<string, unknown>).size as DiagramSize)
  ) {
    throw new Error('The editor received an invalid initialization message.')
  }
  return message as ParentToDialogMessage
}

export function parseDialogMessage(value: string): DialogToParentMessage {
  const message: unknown = JSON.parse(value)
  if (!message || typeof message !== 'object') {
    throw new Error('The editor returned an invalid message.')
  }

  const data = message as Record<string, unknown>
  if (data.type === 'ready' || data.type === 'cancel') {
    return { type: data.type }
  }
  if (
    data.type === 'save' &&
    typeof data.source === 'string' &&
    DIAGRAM_THEMES.includes(data.theme as DiagramTheme) &&
    DIAGRAM_SIZES.includes(data.size as DiagramSize)
  ) {
    return {
      type: 'save',
      source: data.source,
      theme: data.theme as DiagramTheme,
      size: data.size as DiagramSize,
    }
  }

  throw new Error('The editor returned an unsupported message.')
}
