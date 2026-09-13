import { describe, expect, it, vi } from 'vitest'
import { DEFAULT_DIAGRAM_SETTINGS } from './diagramSettings'
import {
  CONTENT_CONTROL_TAG_PREFIX,
  EXCEL_SHAPE_NAME_PREFIX,
  createDiagramPayload,
  getContentControlTag,
  getDiagramIdFromTag,
  getDiagramIdFromExcelShapeName,
  getExcelShapeName,
  parseDiagramPayload,
  sameDiagramPayload,
  type DiagramPayload,
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

  it.each([['png'], ['svg'], null, 1, true, {}, { toString: 'png' }].map(format => ({ format })))(
    'rejects non-string formats: %j',
    ({ format }) => {
      const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')
      expect(() => parseDiagramPayload(JSON.stringify({ ...payload, format })))
        .toThrow('unsupported or invalid structure')
    },
  )

  it.each(['png', 'svg'] as const)('accepts the %s format', (format) => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', format)
    expect(parseDiagramPayload(JSON.stringify(payload))).toEqual(payload)
  })

  it('compares editable payload fields with normalized optional settings', () => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')
    const explicitDefaults = { ...payload, settings: { ...DEFAULT_DIAGRAM_SETTINGS } }
    expect(sameDiagramPayload(payload, { ...payload })).toBe(true)
    expect(sameDiagramPayload(payload, explicitDefaults)).toBe(true)
    expect(sameDiagramPayload(explicitDefaults, payload)).toBe(true)
    expect(sameDiagramPayload(payload, { ...payload, rendererVersion: 'older-renderer' })).toBe(true)
    expect(sameDiagramPayload(explicitDefaults, {
      ...explicitDefaults,
      settings: { ...DEFAULT_DIAGRAM_SETTINGS, fontSize: 18 },
    })).toBe(false)
  })

  it.each<Partial<DiagramPayload>>([
    { id: 'different-id' },
    { source: 'flowchart LR\nB --> C' },
    { theme: 'dark' },
    { size: 'large' },
    { format: 'svg' },
    { settings: { ...DEFAULT_DIAGRAM_SETTINGS, imageQuality: 'high' } },
  ])('detects a changed payload field: %j', (change) => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')
    expect(sameDiagramPayload(payload, { ...payload, ...change })).toBe(false)
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
