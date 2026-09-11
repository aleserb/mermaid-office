import { describe, expect, it, vi } from 'vitest'
import { createDiagramPayload } from './payload'
import { embedPayloadInSvg, readPayloadFromSvg } from './svgMetadata'

describe('SVG diagram metadata', () => {
  it('round-trips Mermaid source', () => {
    vi.stubGlobal('crypto', { randomUUID: () => 'svg-id' })
    const payload = createDiagramPayload('sequenceDiagram\nA->>B: Hello', 'svg')
    const svg = embedPayloadInSvg(
      '<svg xmlns="http://www.w3.org/2000/svg"><text>Hello</text></svg>',
      payload,
    )

    expect(svg).toContain('<metadata id="mermaid-office">')
    expect(readPayloadFromSvg(svg)).toEqual(payload)
    vi.unstubAllGlobals()
  })
})
