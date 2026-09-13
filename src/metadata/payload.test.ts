import { describe, expect, it, vi } from 'vitest'
import {
  CONTENT_CONTROL_TAG_PREFIX,
  EXCEL_SHAPE_NAME_PREFIX,
  createDiagramPayload,
  getContentControlTag,
  getDiagramIdFromTag,
  getDiagramIdFromExcelShapeName,
  getExcelShapeName,
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
      size: 'medium',
    })
    expect(getContentControlTag(payload.id)).toBe(`${CONTENT_CONTROL_TAG_PREFIX}diagram-id`)
    expect(getDiagramIdFromTag(getContentControlTag(payload.id))).toBe('diagram-id')
    expect(getDiagramIdFromTag('other:add-in')).toBeNull()
    expect(getExcelShapeName(payload.id)).toBe(`${EXCEL_SHAPE_NAME_PREFIX}diagram-id`)
    expect(getDiagramIdFromExcelShapeName(getExcelShapeName(payload.id))).toBe('diagram-id')
    expect(getDiagramIdFromExcelShapeName('ordinary image')).toBeNull()

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

  it('defaults diagrams saved before size presets to medium', () => {
    const payload = {
      schemaVersion: 1,
      id: 'old-diagram',
      source: 'flowchart LR\nA --> B',
      theme: 'default',
      format: 'png',
      rendererVersion: '11.17.2',
    }

    expect(parseDiagramPayload(JSON.stringify(payload)).size).toBe('medium')
  })
})
