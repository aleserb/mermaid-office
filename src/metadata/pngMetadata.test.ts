import { describe, expect, it, vi } from 'vitest'
import { base64ToBytes, bytesToBase64 } from './base64'
import { createDiagramPayload } from './payload'
import { embedPayloadInPng, readPayloadFromPng } from './pngMetadata'

const transparentPixel =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+XwX9WQAAAABJRU5ErkJggg=='

describe('PNG diagram metadata', () => {
  it('round-trips Unicode Mermaid source in an iTXt chunk', () => {
    vi.stubGlobal('crypto', { randomUUID: () => 'png-id' })
    const payload = createDiagramPayload('flowchart LR\nA[Привет] --> B', 'png')
    const png = embedPayloadInPng(transparentPixel, payload)

    expect(readPayloadFromPng(png)).toEqual(payload)
    vi.unstubAllGlobals()
  })

  it('rejects a corrupted metadata chunk', () => {
    vi.stubGlobal('crypto', { randomUUID: () => 'png-id' })
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')
    const bytes = base64ToBytes(embedPayloadInPng(transparentPixel, payload))
    bytes[bytes.length - 20] ^= 1

    expect(() => readPayloadFromPng(bytesToBase64(bytes))).toThrow('integrity check')
    vi.unstubAllGlobals()
  })
})
