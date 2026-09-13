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

function metadataChunk(source: string): Uint8Array {
  const bytes = base64ToBytes(embedPayloadInPng(
    transparentPixel,
    createDiagramPayload(source, 'png'),
  ))
  return bytes.slice(base64ToBytes(transparentPixel).length - 12, -12)
}

function appendChunks(...chunks: Uint8Array[]): string {
  const png = base64ToBytes(transparentPixel)
  return bytesToBase64(new Uint8Array([
    ...png.subarray(0, -12),
    ...chunks.flatMap(chunk => Array.from(chunk)),
    ...png.subarray(-12),
  ]))
}

function textChunk(type: 'tEXt' | 'iTXt', keyword: string): Uint8Array {
  const data = new TextEncoder().encode(`${keyword}${type === 'iTXt' ? '\0\0\0\0\0' : '\0'}unrelated text`)
  const chunk = new Uint8Array(data.length + 12)
  const view = new DataView(chunk.buffer)
  view.setUint32(0, data.length)
  chunk.set(new TextEncoder().encode(type), 4)
  chunk.set(data, 8)
  let crc = 0xffffffff
  for (const byte of chunk.subarray(4, -4)) {
    crc ^= byte
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0)
    }
  }
  view.setUint32(chunk.length - 4, (crc ^ 0xffffffff) >>> 0)
  return chunk
}

describe('PNG diagram metadata', () => {
  it('replaces existing payloads without growing the PNG on repeated embeds', () => {
    const original = createDiagramPayload('flowchart LR\nA --> B', 'png')
    const updated = { ...original, source: 'flowchart LR\nB --> C', theme: 'dark' as const }
    const first = embedPayloadInPng(transparentPixel, original)
    const second = embedPayloadInPng(first, updated)

    expect(readPayloadFromPng(second)).toEqual(updated)
    expect(second).toBe(embedPayloadInPng(transparentPixel, updated))
    expect(embedPayloadInPng(second, updated)).toBe(second)
  })

  it('reads the last legacy duplicate and removes all duplicates when embedding', () => {
    const first = metadataChunk('flowchart LR\nA --> B')
    const last = metadataChunk('flowchart LR\nB --> C')
    const legacy = appendChunks(first, last)
    const updated = createDiagramPayload('flowchart LR\nC --> D', 'png')

    expect(readPayloadFromPng(legacy)?.source).toBe('flowchart LR\nB --> C')
    expect(embedPayloadInPng(legacy, updated))
      .toBe(embedPayloadInPng(transparentPixel, updated))
  })

  it('still rejects a corrupted later duplicate instead of returning stale metadata', () => {
    const first = metadataChunk('flowchart LR\nA --> B')
    const last = metadataChunk('flowchart LR\nB --> C')
    last[last.length - 1] ^= 1
    expect(() => readPayloadFromPng(appendChunks(first, last))).toThrow('integrity check')
  })

  it('preserves unrelated chunks and physical density when replacing metadata', () => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')
    const withDensity = setPngPhysicalWidth(transparentPixel, 100)
    const embedded = embedPayloadInPng(withDensity, payload)
    const updated = { ...payload, source: 'flowchart LR\nB --> C' }
    expect(embedPayloadInPng(embedded, updated)).toBe(embedPayloadInPng(withDensity, updated))
  })

  it('preserves unrelated text chunks but replaces every reserved-key text chunk', () => {
    const unrelatedText = textChunk('tEXt', 'Author')
    const unrelatedInternationalText = textChunk('iTXt', 'mermaid-office-other')
    const original = appendChunks(
      unrelatedText,
      metadataChunk('flowchart LR\nA --> B'),
      textChunk('tEXt', 'mermaid-office'),
      unrelatedInternationalText,
    )
    const payload = createDiagramPayload('flowchart LR\nB --> C', 'png')
    expect(embedPayloadInPng(original, payload)).toBe(embedPayloadInPng(
      appendChunks(unrelatedText, unrelatedInternationalText),
      payload,
    ))
  })

  it('rejects truncated chunks while embedding rather than copying invalid boundaries', () => {
    const bytes = base64ToBytes(transparentPixel)
    new DataView(bytes.buffer).setUint32(8, bytes.length)
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')
    expect(() => embedPayloadInPng(bytesToBase64(bytes), payload)).toThrow('truncated chunk')
  })

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
