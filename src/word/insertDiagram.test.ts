import { afterEach, describe, expect, it, vi } from 'vitest'
import { isSvgInsertionSupported } from './insertDiagram'

describe('isSvgInsertionSupported', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
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
