import { describe, expect, it, vi } from 'vitest'
import { textToBase64 } from './base64'
import { readPayloadFromImage } from './imageMetadata'
import { createDiagramPayload } from './payload'
import { embedPayloadInPng } from './pngMetadata'
import { embedPayloadInSvg } from './svgMetadata'

const transparentPixel =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+Xv6X9QAAAABJRU5ErkJggg=='

describe('image metadata', () => {
  it('reads Mermaid payloads from PNG image data', () => {
    vi.stubGlobal('crypto', { randomUUID: () => 'png-image-id' })
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')

    expect(readPayloadFromImage(embedPayloadInPng(transparentPixel, payload))).toEqual(
      payload,
    )
    vi.unstubAllGlobals()
  })

  it('reads Mermaid payloads from SVG image data URLs', () => {
    vi.stubGlobal('crypto', { randomUUID: () => 'svg-image-id' })
    const payload = createDiagramPayload('sequenceDiagram\nA->>B: Hello', 'svg')
    const svg = embedPayloadInSvg(
      '<svg xmlns="http://www.w3.org/2000/svg"><text>Hello</text></svg>',
      payload,
    )

    expect(
      readPayloadFromImage(`data:image/svg+xml;base64,${textToBase64(svg)}`),
    ).toEqual(payload)
    vi.unstubAllGlobals()
  })

  it('ignores images without Mermaid metadata', () => {
    expect(readPayloadFromImage(textToBase64('<svg xmlns="http://www.w3.org/2000/svg"/>'))).toBeNull()
  })
})
