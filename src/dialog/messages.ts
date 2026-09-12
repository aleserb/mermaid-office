import {
  DIAGRAM_SIZES,
  DIAGRAM_THEMES,
  type DiagramSize,
  type DiagramTheme,
} from '../metadata/payload'

export type ParentToDialogMessage = {
  type: 'initialize'
  mode: 'insert' | 'update'
  source: string
  theme: DiagramTheme
  size: DiagramSize
}

export type DialogToParentMessage =
  | { type: 'ready' }
  | {
      type: 'save'
      source: string
      theme: DiagramTheme
      size: DiagramSize
      svg: string
      raster: { base64: string; width: number; height: number }
    }
  | { type: 'cancel' }

export function parseParentMessage(value: string): ParentToDialogMessage {
  const message: unknown = JSON.parse(value)
  if (
    !message ||
    typeof message !== 'object' ||
    (message as Record<string, unknown>).type !== 'initialize' ||
    !['insert', 'update'].includes((message as Record<string, unknown>).mode as string) ||
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
    DIAGRAM_SIZES.includes(data.size as DiagramSize) &&
    typeof data.svg === 'string' &&
    data.raster &&
    typeof data.raster === 'object' &&
    typeof (data.raster as Record<string, unknown>).base64 === 'string' &&
    typeof (data.raster as Record<string, unknown>).width === 'number' &&
    Number.isFinite((data.raster as Record<string, unknown>).width) &&
    (data.raster as Record<string, number>).width > 0 &&
    typeof (data.raster as Record<string, unknown>).height === 'number' &&
    Number.isFinite((data.raster as Record<string, unknown>).height) &&
    (data.raster as Record<string, number>).height > 0
  ) {
    return {
      type: 'save',
      source: data.source,
      theme: data.theme as DiagramTheme,
      size: data.size as DiagramSize,
      svg: data.svg,
      raster: data.raster as {
        base64: string
        width: number
        height: number
      },
    }
  }

  throw new Error('The editor returned an unsupported message.')
}
