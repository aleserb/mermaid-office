import { afterEach, describe, expect, it, vi } from 'vitest'
import { getRasterDimensions, rasterizeSvg } from './insertDiagram'

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

describe('adaptive raster sizing', () => {
  it('offers smaller Standard and higher-detail High exports without changing the logical frame', () => {
    const width = 1853.84765625
    const height = 3236.24267578125
    expect(getRasterDimensions(width, height, 'medium', 'standard')).toEqual({ width: 864, height: 1508 })
    expect(getRasterDimensions(width, height, 'medium', 'auto')).toEqual({ width: 1550, height: 2705 })
    expect(getRasterDimensions(width, height, 'medium', 'high')).toEqual({ width: 4384, height: 7653 })
  })

  it.each(['auto', 'standard', 'high'] as const)('bounds resource use at %s quality', quality => {
    const limit = quality === 'high' ? { dimension: 8192, pixels: 32 * 1024 * 1024 }
      : { dimension: 4096, pixels: 4 * 1024 * 1024 }
    for (const [width, height] of [[10000, 10000], [10000, 1], [1, 10000]]) {
      const result = getRasterDimensions(width, height, 468, quality)
      expect(result.width).toBeLessThanOrEqual(limit.dimension)
      expect(result.height).toBeLessThanOrEqual(limit.dimension)
      expect(result.width * result.height).toBeLessThanOrEqual(limit.pixels)
    }
  })

  it('targets normal viewing at 2x screen density for small diagrams', () => {
    expect(getRasterDimensions(100, 50, 'medium')).toEqual({ width: 864, height: 432 })
    expect(getRasterDimensions(100, 50, 'page-width')).toEqual({ width: 1248, height: 624 })
  })

  it('preserves native detail for dense diagrams when within budget', () => {
    expect(getRasterDimensions(1200, 2400, 'medium')).toEqual({ width: 1200, height: 2400 })
  })

  it('does not depend on a Word size preset for native-resolution exports', () => {
    expect(getRasterDimensions(300, 100)).toEqual({ width: 300, height: 100 })
  })

  it('adapts to a manually resized Word picture in points', () => {
    expect(getRasterDimensions(100, 50, 216)).toEqual({ width: 576, height: 288 })
    expect(getRasterDimensions(100, 50, 432)).toEqual({ width: 1152, height: 576 })
  })

  it('uses the fitted height limit for new tall diagrams, but not existing widths', () => {
    expect(getRasterDimensions(100, 1000, 'medium')).toEqual({ width: 173, height: 1733 })
    expect(getRasterDimensions(100, 1000, 324)).toEqual({ width: 409, height: 4096 })
  })

  it.each([
    [1853.84765625, 3236.24267578125],
    [10000, 1000],
    [1000, 10000],
    [10000, 10000],
  ])('bounds dimensions and memory for a %s by %s diagram', (width, height) => {
    const result = getRasterDimensions(width, height, 'medium')
    expect(result.width).toBeLessThanOrEqual(4096)
    expect(result.height).toBeLessThanOrEqual(4096)
    expect(result.width * result.height).toBeLessThanOrEqual(4 * 1024 * 1024)
    expect(Math.abs(result.width - result.height * width / height))
      .toBeLessThanOrEqual(1 + width / height)
  })

  it('uses at least 87% fewer pixels for the large sequence reproduction', () => {
    const result = getRasterDimensions(1853.84765625, 3236.24267578125, 'medium')
    expect(result).toEqual({ width: 1550, height: 2705 })
    expect(result.width * result.height).toBeLessThan(0.13 * 4384 * 7653)
    expect(result.width).toBeGreaterThan(324 * (96 / 72) * 2)
  })

  it.each([0, -1, NaN, Infinity])('rejects invalid display width %s', (width) => {
    expect(() => getRasterDimensions(100, 50, width)).toThrow('display width')
  })

  it.each([[0, 10], [10, -1], [NaN, 10], [10, Infinity]])(
    'rejects invalid dimensions %s by %s',
    (width, height) => {
      expect(() => getRasterDimensions(width, height)).toThrow('positive finite numbers')
    },
  )
})

const rasterCases = [
  { label: 'full', left: 0, top: 0, right: 200, bottom: 100, x: 0, y: 0, width: 100, height: 50 },
  { label: 'cropped', left: 20, top: 10, right: 180, bottom: 90, x: 6, y: 1, width: 88, height: 48 },
].flatMap(bounds => ['success', 'unsupported', 'empty', 'error'].flatMap(encoding =>
  (['medium', 216] as const).flatMap(size =>
    (['auto', 'standard', 'high'] as const).map(quality => ({ ...bounds, encoding, size, quality }))),
))

it.each(rasterCases)('renders $label SVG bounds at $size/$quality and releases both canvases on $encoding', async (bounds) => {
  const { encoding } = bounds
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
    toDataURL: vi.fn().mockImplementation(() => {
      if (encoding === 'error') throw new Error('PNG encoding failed')
      if (encoding === 'unsupported') return 'data:,'
      if (encoding === 'empty') return 'data:image/png;base64,'
      return 'data:image/png;base64,rendered-png'
    }),
  })
  vi.spyOn(document, 'createElement').mockReturnValueOnce(probe).mockReturnValueOnce(output)
  class VectorImage {
    src = ''
    decode = vi.fn().mockResolvedValue(undefined)
  }
  vi.stubGlobal('Image', VectorImage)
  const revokeObjectURL = vi.fn()
  vi.stubGlobal('URL', { createObjectURL: () => 'blob:diagram', revokeObjectURL })

  const result = rasterizeSvg('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50"/>', bounds.size, bounds.quality)
  const dimensions = getRasterDimensions(bounds.width, bounds.height, bounds.size, bounds.quality)
  if (encoding === 'success') {
    await expect(result).resolves.toEqual({ base64: 'rendered-png', width: bounds.width, height: bounds.height })
  } else {
    await expect(result).rejects.toThrow(encoding === 'error' ? 'PNG encoding failed' : 'PNG resolution')
  }
  expect(output.width).toBe(1)
  expect(output.height).toBe(1)
  expect(outputContext.drawImage).toHaveBeenCalledExactlyOnceWith(
    expect.any(VectorImage), bounds.x, bounds.y, bounds.width, bounds.height,
    0, 0, dimensions.width, dimensions.height,
  )
  expect(probe.width).toBe(1)
  expect(probe.height).toBe(1)
  expect(revokeObjectURL).toHaveBeenCalledExactlyOnceWith('blob:diagram')
})
