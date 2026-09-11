import { afterEach, describe, expect, it, vi } from 'vitest'
import { fitDiagram, isSvgInsertionSupported } from './insertDiagram'

describe('isSvgInsertionSupported', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  describe('fitDiagram', () => {
    it('keeps small diagrams at their natural size', () => {
      expect(fitDiagram(200, 100)).toEqual({ width: 200, height: 100 })
    })

    it('scales large diagrams without changing their aspect ratio', () => {
      expect(fitDiagram(1000, 500)).toEqual({ width: 500, height: 250 })
      expect(fitDiagram(500, 1300)).toEqual({ width: 250, height: 650 })
    })
  })

  it('returns false outside an Office host', () => {
    vi.stubGlobal('Office', undefined)
    expect(isSvgInsertionSupported()).toBe(false)
  })

  it('checks for ImageCoercion 1.2', () => {
    const isSetSupported = vi.fn().mockReturnValue(true)
    vi.stubGlobal('Office', {
      context: { requirements: { isSetSupported } },
    })

    expect(isSvgInsertionSupported()).toBe(true)
    expect(isSetSupported).toHaveBeenCalledWith('ImageCoercion', '1.2')
  })
})
