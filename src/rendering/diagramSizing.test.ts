import { describe, expect, it } from 'vitest'
import { DIAGRAM_WIDTHS, fitDiagram } from './diagramSizing'

describe('fitDiagram', () => {
  it('keeps small diagrams at their natural size', () => {
    expect(fitDiagram(200, 100)).toEqual({ width: 200, height: 100 })
  })

  it('scales large diagrams without changing their aspect ratio', () => {
    expect(fitDiagram(1000, 500)).toEqual({ width: 500, height: 250 })
    expect(fitDiagram(500, 1300)).toEqual({ width: 250, height: 650 })
  })

  it('can upscale a diagram to a selected width preset', () => {
    expect(fitDiagram(200, 100, 400, 650, true)).toEqual({
      width: 400,
      height: 200,
    })
  })

  it.each([
    ['small', 216],
    ['medium', 324],
    ['large', 396],
    ['page-width', 468],
  ] as const)('uses the shared %s preset in points', (size, width) => {
    expect(DIAGRAM_WIDTHS[size]).toBe(width)
    expect(fitDiagram(100, 50, DIAGRAM_WIDTHS[size], 650, true))
      .toEqual({ width, height: width / 2 })
  })
})
