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

  it.each([
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<!-- diagram exported from Office -->',
    '<?xml version="1.0"?>\n<!-- exported diagram -->\n',
  ])('recognizes SVG after an XML prolog: %s', (prefix) => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'svg')
    const svg = embedPayloadInSvg('<svg xmlns="http://www.w3.org/2000/svg"/>', payload)
    expect(readPayloadFromImage(textToBase64(`${prefix}${svg}`))).toEqual(payload)
  })

  it('recognizes a namespace-prefixed SVG root', () => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'svg')
    const svg = embedPayloadInSvg(
      '<s:svg xmlns:s="http://www.w3.org/2000/svg"/>',
      payload,
    )
    expect(readPayloadFromImage(textToBase64(svg))).toEqual(payload)
  })

  it.each([
    '<svg>',
    '<svg xmlns="urn:not-svg">',
    '<SVG xmlns="http://www.w3.org/2000/svg">',
    '<document xmlns="http://www.w3.org/2000/svg">',
  ])('ignores metadata under a non-SVG root: %s', (root) => {
    const payload = createDiagramPayload('flowchart LR\nA --> B', 'svg')
    const rootName = root.slice(1).split(/[\s>]/)[0]
    const xml = `${root}<metadata id="mermaid-office">${textToBase64(JSON.stringify(payload))}</metadata></${rootName}>`
    expect(readPayloadFromImage(textToBase64(xml))).toBeNull()
  })

  it.each(['<broken>', '<!-- unfinished', '<svg/>'])(
    'does not read metadata from malformed XML ending with %s',
    (suffix) => {
      const payload = createDiagramPayload('flowchart LR\nA --> B', 'svg')
      const svg = embedPayloadInSvg('<svg xmlns="http://www.w3.org/2000/svg"/>', payload)
      expect(readPayloadFromImage(textToBase64(`${svg}${suffix}`))).toBeNull()
    },
  )

  it('preserves metadata errors for SVGs with an XML declaration', () => {
    const svg = `<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg"><metadata id="mermaid-office">${textToBase64('not JSON')}</metadata></svg>`
    expect(() => readPayloadFromImage(textToBase64(svg))).toThrow('not valid JSON')
  })
})
