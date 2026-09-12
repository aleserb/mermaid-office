import { describe, expect, it, vi } from 'vitest'
import { base64ToBytes, bytesToBase64 } from './base64'
import { createDiagramPayload } from './payload'
import {
  embedPayloadInPng,
  getPngDimensions,
  readPayloadFromPng,
  setPngPhysicalWidth,
} from './pngMetadata'

const transparentPixel =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+XwX9WQAAAABJRU5ErkJggg=='

describe('PNG diagram metadata', () => {
  it('reads pixel dimensions independently of embedded source and physical density', () => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')
    expect(getPngDimensions(transparentPixel)).toEqual({ width: 1, height: 1 })
    expect(getPngDimensions(setPngPhysicalWidth(embedPayloadInPng(transparentPixel, payload), 324)))
      .toEqual({ width: 1, height: 1 })
  })

  it.each(['missing', 'truncated', 'invalid length', 'zero width', 'zero height'])(
    'rejects a PNG with a %s header',
    (scenario) => {
      let bytes = base64ToBytes(transparentPixel)
      const view = new DataView(bytes.buffer)
      if (scenario === 'missing') bytes = bytes.subarray(0, 8)
      if (scenario === 'truncated') bytes = bytes.subarray(0, 24)
      if (scenario === 'invalid length') view.setUint32(8, 1)
      if (scenario === 'zero width') view.setUint32(16, 0)
      if (scenario === 'zero height') view.setUint32(20, 0)
      expect(() => getPngDimensions(bytesToBase64(bytes))).toThrow(/IHDR|dimensions/)
    },
  )

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

  it('sets PNG density from the intended Word width', () => {
    const bytes = base64ToBytes(setPngPhysicalWidth(transparentPixel, 0.375))
    const type = new TextDecoder().decode(bytes.subarray(37, 41))
    const density = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)

    expect(type).toBe('pHYs')
    expect(density.getUint32(41)).toBe(7559)
    expect(density.getUint32(45)).toBe(7559)
    expect(bytes[49]).toBe(1)
  })
})
