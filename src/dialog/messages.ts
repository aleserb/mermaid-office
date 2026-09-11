export type ParentToDialogMessage = {
  type: 'initialize'
  source: string
}

export type DialogToParentMessage =
  | { type: 'ready' }
  | { type: 'save'; source: string }
  | { type: 'cancel' }

export function parseParentMessage(value: string): ParentToDialogMessage {
  const message: unknown = JSON.parse(value)
  if (
    !message ||
    typeof message !== 'object' ||
    (message as Record<string, unknown>).type !== 'initialize' ||
    typeof (message as Record<string, unknown>).source !== 'string'
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
  if (data.type === 'save' && typeof data.source === 'string') {
    return { type: 'save', source: data.source }
  }

  throw new Error('The editor returned an unsupported message.')
}
