import { describe, expect, it } from 'vitest'
import { DEFAULT_DIAGRAM_SETTINGS, getDiagramSettings, isDiagramSettings, sameDiagramSettings } from './diagramSettings'
import { createDiagramPayload, parseDiagramPayload } from './payload'
import { embedPayloadInPng, readPayloadFromPng } from './pngMetadata'

const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+XwX9WQAAAABJRU5ErkJggg=='

describe('diagram settings metadata', () => {
  it('uses unchanged defaults for legacy diagrams without settings', () => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')
    expect(parseDiagramPayload(JSON.stringify(payload))).not.toHaveProperty('settings')
    expect(getDiagramSettings(payload.settings)).toEqual(DEFAULT_DIAGRAM_SETTINGS)
    expect(sameDiagramSettings(undefined, { ...DEFAULT_DIAGRAM_SETTINGS })).toBe(true)
  })

  it('preserves all settings in document and PNG metadata without sharing mutable input', () => {
    const settings = {
      ...DEFAULT_DIAGRAM_SETTINGS, fontSize: 20 as const, look: 'handDrawn' as const,
      spacing: 'compact' as const, curve: 'step' as const, layout: 'elk' as const,
      sequenceNumbers: true, sequenceWrap: true, sequenceMirrorActors: false,
      imageQuality: 'high' as const,
    }
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png', 'forest', 'medium', settings)
    expect(parseDiagramPayload(JSON.stringify(payload)).settings).toEqual(settings)
    expect(readPayloadFromPng(embedPayloadInPng(png, payload))?.settings).toEqual(settings)
    settings.sequenceNumbers = false
    expect(payload.settings?.sequenceNumbers).toBe(true)
  })

  it.each([
    { fontFamily: 'url(https://example.com/font)' }, { fontFamily: {} },
    { fontSize: '18' }, { fontSize: -1 }, { fontSize: 999 },
    { look: 'unknown' }, { spacing: null }, { curve: [] }, { layout: {} },
    { imageQuality: 'unlimited' }, { sequenceNumbers: 'true' },
    { sequenceWrap: 1 }, { sequenceMirrorActors: undefined },
  ])('rejects unsupported settings %j', (invalid) => {
    const settings = { ...DEFAULT_DIAGRAM_SETTINGS, ...invalid }
    expect(isDiagramSettings(settings)).toBe(false)
    const payload = { ...createDiagramPayload('flowchart LR\nA', 'png'), settings }
    expect(() => parseDiagramPayload(JSON.stringify(payload))).toThrow('unsupported or invalid')
  })

  it('compares settings by their values, including quality and sequence toggles', () => {
    expect(sameDiagramSettings(DEFAULT_DIAGRAM_SETTINGS, { ...DEFAULT_DIAGRAM_SETTINGS })).toBe(true)
    expect(sameDiagramSettings(undefined, { ...DEFAULT_DIAGRAM_SETTINGS, imageQuality: 'high' })).toBe(false)
    expect(sameDiagramSettings(undefined, { ...DEFAULT_DIAGRAM_SETTINGS, sequenceWrap: true })).toBe(false)
  })
})
