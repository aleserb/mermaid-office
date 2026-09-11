import { describe, expect, it, vi } from 'vitest'
import {
  CONTENT_CONTROL_TAG_PREFIX,
  createDiagramPayload,
  getContentControlTag,
  getDiagramIdFromTag,
  parseDiagramPayload,
} from './payload'

describe('diagram payload', () => {
  it('creates a versioned payload and content-control tag', () => {
    vi.stubGlobal('crypto', { randomUUID: () => 'diagram-id' })

    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')

    expect(payload).toMatchObject({
      schemaVersion: 1,
      id: 'diagram-id',
      format: 'png',
      theme: 'default',
    })
    expect(getContentControlTag(payload.id)).toBe(`${CONTENT_CONTROL_TAG_PREFIX}diagram-id`)
    expect(getDiagramIdFromTag(getContentControlTag(payload.id))).toBe('diagram-id')
    expect(getDiagramIdFromTag('other:add-in')).toBeNull()

    vi.unstubAllGlobals()
  })

  it('rejects unsupported metadata', () => {
    expect(() => parseDiagramPayload('{"schemaVersion":2}')).toThrow(
      'unsupported or invalid structure',
    )
  })

  it('accepts the Redux Color theme in stored diagram metadata', () => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png', 'redux-color')

    expect(parseDiagramPayload(JSON.stringify(payload)).theme).toBe('redux-color')
  })
})
