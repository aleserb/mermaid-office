import { describe, expect, it } from 'vitest'
import { createDiagramPictureOoxml } from './pictureOoxml'
import { createDiagramPayload } from '../metadata/payload'
import { embedPayloadInPng, readPayloadFromPng, setPngPhysicalWidth } from '../metadata/pngMetadata'

const transparentPixel =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+XwX9WQAAAABJRU5ErkJggg=='
const pkg = 'http://schemas.microsoft.com/office/2006/xmlPackage'
const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const wordDrawing = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships'

function parse(xml: string) {
  const document = new DOMParser().parseFromString(xml, 'application/xml')
  expect(document.getElementsByTagName('parsererror')).toHaveLength(0)
  return document
}

describe('Word PNG drawing package', () => {
  it.each([
    [324, 565.6034482758621],
    [267.75, 467.35],
    [216, 108],
  ])('fits the bitmap to the complete %s x %s point frame', (width, height) => {
    const document = parse(createDiagramPictureOoxml(transparentPixel, { width, height }))
    const outer = document.getElementsByTagNameNS(wordDrawing, 'extent')[0]
    const inner = document.getElementsByTagNameNS(drawing, 'ext')[1]
    for (const extent of [outer, inner]) {
      expect(extent.getAttribute('cx')).toBe(String(Math.round(width * 12700)))
      expect(extent.getAttribute('cy')).toBe(String(Math.round(height * 12700)))
    }
    const origin = document.getElementsByTagNameNS(drawing, 'off')[0]
    expect(origin.getAttribute('x')).toBe('0')
    expect(origin.getAttribute('y')).toBe('0')
    expect(document.getElementsByTagNameNS(drawing, 'stretch')).toHaveLength(1)
    expect(document.getElementsByTagNameNS(drawing, 'fillRect')).toHaveLength(1)
    expect(document.getElementsByTagNameNS(drawing, 'srcRect')).toHaveLength(0)
    expect(document.getElementsByTagNameNS('*', 'useLocalDpi')[0].getAttribute('val')).toBe('0')
  })

  it('contains one inline picture and no content controls or section formatting', () => {
    const document = parse(createDiagramPictureOoxml(transparentPixel, { width: 324, height: 200 }))
    expect(document.getElementsByTagNameNS(word, 'p')).toHaveLength(1)
    expect(document.getElementsByTagNameNS(word, 'drawing')).toHaveLength(1)
    expect(document.getElementsByTagNameNS(wordDrawing, 'inline')).toHaveLength(1)
    expect(document.getElementsByTagNameNS(word, 'sdt')).toHaveLength(0)
    expect(document.getElementsByTagNameNS(word, 'sectPr')).toHaveLength(0)
    const parts = Array.from(document.getElementsByTagNameNS(pkg, 'part'))
    expect(parts.map(part => part.getAttributeNS(pkg, 'name'))).toEqual([
      '/_rels/.rels', '/word/document.xml', '/word/_rels/document.xml.rels', '/word/media/diagram.png',
    ])
    const links = document.getElementsByTagNameNS(relationships, 'Relationship')
    expect(links).toHaveLength(2)
    expect(links[0].getAttribute('Target')).toBe('word/document.xml')
    expect(links[1].getAttribute('Target')).toBe('media/diagram.png')
    expect(document.getElementsByTagNameNS(drawing, 'blip')[0].getAttributeNS(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed',
    )).toBe(links[1].getAttribute('Id'))
    expect(parts[3].getAttributeNS(pkg, 'contentType')).toBe('image/png')
  })

  it('embeds the original PNG bytes, source payload, and physical density unchanged', () => {
    const payload = createDiagramPayload('sequenceDiagram\nA->>B: Message', 'png', 'redux-color')
    const png = setPngPhysicalWidth(embedPayloadInPng(transparentPixel, payload), 324)
    const document = parse(createDiagramPictureOoxml(png, { width: 324, height: 500 }))
    const embedded = document.getElementsByTagNameNS(pkg, 'binaryData')[0].textContent!
    expect(embedded).toBe(png)
    expect(readPayloadFromPng(embedded)).toEqual(payload)
  })

  it.each([
    ['Custom "title" & <diagram>', 'Description with </wp:docPr> & "quotes"'],
    ['', ''],
  ])('preserves alt text without interpreting XML characters', (altTextTitle, altTextDescription) => {
    const document = parse(createDiagramPictureOoxml(transparentPixel, {
      width: 324, height: 200, altTextTitle, altTextDescription,
    }))
    const properties = document.getElementsByTagNameNS(wordDrawing, 'docPr')
    expect(properties).toHaveLength(1)
    expect(properties[0].getAttribute('title')).toBe(altTextTitle)
    expect(properties[0].getAttribute('descr')).toBe(altTextDescription)
  })

  it('provides default accessible text for newly inserted diagrams', () => {
    const document = parse(createDiagramPictureOoxml(transparentPixel, { width: 324, height: 200 }))
    const properties = document.getElementsByTagNameNS(wordDrawing, 'docPr')[0]
    expect(properties.getAttribute('title')).toBe('Mermaid diagram')
    expect(properties.getAttribute('descr')).toBe('Diagram created with Mermaid Office.')
  })

  it.each([0, -1, Number.NaN, Number.POSITIVE_INFINITY, Number.MAX_VALUE, 0.000001])(
    'rejects unrepresentable Word dimensions: %s', (invalid) => {
      expect(() => createDiagramPictureOoxml(transparentPixel, { width: invalid, height: 100 }))
        .toThrow('Diagram dimensions')
      expect(() => createDiagramPictureOoxml(transparentPixel, { width: 100, height: invalid }))
        .toThrow('Diagram dimensions')
    },
  )
})
