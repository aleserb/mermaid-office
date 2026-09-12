import { afterEach, describe, expect, it, vi } from 'vitest'
import { getRasterDimensions, rasterizeSvg } from './insertDiagram'

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

describe('high-detail raster sizing', () => {
  it('uses four pixels per displayed CSS pixel for small diagrams', () => {
    expect(getRasterDimensions(100, 50, 'medium')).toEqual({ width: 1728, height: 864 })
    expect(getRasterDimensions(100, 50, 'page-width')).toEqual({ width: 2496, height: 1248 })
  })

  it('preserves twice the native detail for dense diagrams when within budget', () => {
    expect(getRasterDimensions(1600, 2400, 'medium')).toEqual({ width: 3200, height: 4800 })
  })

  it('does not depend on a Word size preset for native-resolution exports', () => {
    expect(getRasterDimensions(300, 100)).toEqual({ width: 600, height: 200 })
  })

  it.each([
    [1853.84765625, 3236.24267578125],
    [10000, 1000],
    [1000, 10000],
    [10000, 10000],
  ])('bounds dimensions and memory for a %s by %s diagram', (width, height) => {
    const result = getRasterDimensions(width, height, 'medium')
    expect(result.width).toBeLessThanOrEqual(8192)
    expect(result.height).toBeLessThanOrEqual(8192)
    expect(result.width * result.height).toBeLessThanOrEqual(16 * 1024 * 1024)
    expect(Math.abs(result.width - result.height * width / height))
      .toBeLessThanOrEqual(1 + width / height)
  })

  it('retains substantially more detail for the large sequence reproduction', () => {
    const result = getRasterDimensions(1853.84765625, 3236.24267578125, 'medium')
    expect(result.width).toBeGreaterThan(3000)
    expect(result.height).toBeGreaterThan(5000)
    expect(result.width * result.height).toBeGreaterThan(10 * 864 * 1508)
  })

  it.each([[0, 10], [10, -1], [NaN, 10], [10, Infinity]])(
    'rejects invalid dimensions %s by %s',
    (width, height) => {
      expect(() => getRasterDimensions(width, height)).toThrow('positive finite numbers')
    },
  )
})

it.each([
  { label: 'full', left: 0, top: 0, right: 200, bottom: 100, x: 0, y: 0, width: 100, height: 50 },
  { label: 'cropped', left: 20, top: 10, right: 180, bottom: 90, x: 6, y: 1, width: 88, height: 48 },
])('renders $label SVG bounds directly at final resolution and releases the probe canvas', async (bounds) => {
  const pixels = new Uint8ClampedArray(200 * 100 * 4)
  for (let y = bounds.top; y < bounds.bottom; y += 1) {
    for (let x = bounds.left; x < bounds.right; x += 1) {
      pixels[(y * 200 + x) * 4 + 3] = 255
    }
  }
  const probeContext = {
    scale: vi.fn(),
    drawImage: vi.fn(),
    getImageData: vi.fn().mockReturnValue({ data: pixels }),
  }
  const outputContext = { drawImage: vi.fn() }
  const probe = Object.assign(document.createElement('canvas'), {
    getContext: vi.fn().mockReturnValue(probeContext),
  })
  const output = Object.assign(document.createElement('canvas'), {
    getContext: vi.fn().mockReturnValue(outputContext),
    toDataURL: vi.fn().mockReturnValue('data:image/png;base64,rendered-png'),
  })
  vi.spyOn(document, 'createElement').mockReturnValueOnce(probe).mockReturnValueOnce(output)
  class VectorImage {
    src = ''
    decode = vi.fn().mockResolvedValue(undefined)
  }
  vi.stubGlobal('Image', VectorImage)
  const revokeObjectURL = vi.fn()
  vi.stubGlobal('URL', { createObjectURL: () => 'blob:diagram', revokeObjectURL })

  const result = await rasterizeSvg('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50"/>', 'medium')
  const dimensions = getRasterDimensions(bounds.width, bounds.height, 'medium')
  expect(result).toEqual({ base64: 'rendered-png', width: bounds.width, height: bounds.height })
  expect(output.width).toBe(dimensions.width)
  expect(output.height).toBe(dimensions.height)
  expect(outputContext.drawImage).toHaveBeenCalledExactlyOnceWith(
    expect.any(VectorImage), bounds.x, bounds.y, bounds.width, bounds.height,
    0, 0, dimensions.width, dimensions.height,
  )
  expect(probe.width).toBe(1)
  expect(probe.height).toBe(1)
  expect(revokeObjectURL).toHaveBeenCalledExactlyOnceWith('blob:diagram')
})
