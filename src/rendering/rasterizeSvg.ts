import type { ImageQuality } from '../metadata/diagramSettings'
import type { DiagramSize } from '../metadata/payload'
import { getSvgDimensions } from '../mermaid/svgDimensions'
import { DIAGRAM_WIDTHS, fitDiagram } from './diagramSizing'

export interface RasterizedDiagram {
  base64: string
  width: number
  height: number
}

interface PixelBounds {
  left: number
  top: number
  width: number
  height: number
}

const RASTER_LIMITS = {
  auto: { dimension: 4096, pixels: 4 * 1024 * 1024 },
  standard: { dimension: 4096, pixels: 4 * 1024 * 1024 },
  high: { dimension: 8192, pixels: 32 * 1024 * 1024 },
} satisfies Record<ImageQuality, { dimension: number; pixels: number }>

export function getRasterDimensions(
  width: number,
  height: number,
  sizeOrWidth?: DiagramSize | number,
  quality: ImageQuality = 'auto',
): { width: number; height: number } {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw new Error('Diagram dimensions must be positive finite numbers.')
  }
  if (typeof sizeOrWidth === 'number' && (!Number.isFinite(sizeOrWidth) || sizeOrWidth <= 0)) {
    throw new Error('Diagram display width must be a positive finite number.')
  }
  const displayWidth = typeof sizeOrWidth === 'number'
    ? sizeOrWidth
    : sizeOrWidth ? fitDiagram(width, height, DIAGRAM_WIDTHS[sizeOrWidth], 650, true).width : 0
  // Display widths are in points. Standard targets a 2x-density display; Auto
  // retains native SVG detail, while High reserves detail for extreme zoom.
  const displayScale = displayWidth > 0 ? displayWidth * (96 / 72) * 2 / width : 1
  const preferredScale = quality === 'high'
    ? Math.max(3, displayWidth > 0 ? displayScale * 4 : 3)
    : quality === 'standard' ? displayScale : Math.max(1, displayScale)
  const limits = RASTER_LIMITS[quality]
  const scale = Math.min(
    preferredScale,
    limits.dimension / width,
    limits.dimension / height,
    Math.sqrt(limits.pixels / width / height),
  )
  return {
    width: Math.max(1, Math.floor(width * scale)),
    height: Math.max(1, Math.floor(height * scale)),
  }
}

export function findVisiblePixelBounds(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
): PixelBounds | null {
  let left = width
  let top = height
  let right = -1
  let bottom = -1

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (pixels[(y * width + x) * 4 + 3] === 0) {
        continue
      }
      left = Math.min(left, x)
      top = Math.min(top, y)
      right = Math.max(right, x)
      bottom = Math.max(bottom, y)
    }
  }

  return right < left
    ? null
    : {
        left,
        top,
        width: right - left + 1,
        height: bottom - top + 1,
      }
}

export function normalizeSvgDimensions(svg: string): {
  svg: string
  width: number
  height: number
} {
  const svgDocument = new DOMParser().parseFromString(svg, 'image/svg+xml')
  const root = svgDocument.documentElement
  const { width: sourceWidth, height: sourceHeight } = getSvgDimensions(root)
  root.setAttribute('width', String(sourceWidth))
  root.setAttribute('height', String(sourceHeight))
  root.style.removeProperty('max-width')

  return {
    svg: new XMLSerializer().serializeToString(root),
    width: sourceWidth,
    height: sourceHeight,
  }
}

export async function rasterizeSvg(
  svg: string,
  sizeOrWidth?: DiagramSize | number,
  quality: ImageQuality = 'auto',
): Promise<RasterizedDiagram> {
  const normalized = normalizeSvgDimensions(svg)
  const scale = Math.min(2, 4096 / Math.max(normalized.width, normalized.height))
  const width = Math.max(1, Math.round(normalized.width * scale))
  const height = Math.max(1, Math.round(normalized.height * scale))
  const url = URL.createObjectURL(new Blob([normalized.svg], { type: 'image/svg+xml' }))

  try {
    const image = new Image()
    image.src = url
    await image.decode()

    const canvas = window.document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const context = canvas.getContext('2d')
    if (!context) {
      throw new Error('This browser cannot create the PNG fallback.')
    }

    context.scale(scale, scale)
    context.drawImage(image, 0, 0, normalized.width, normalized.height)

    const bounds = findVisiblePixelBounds(
      context.getImageData(0, 0, width, height).data,
      width,
      height,
    )
    if (!bounds) {
      throw new Error('The rendered diagram does not contain any visible content.')
    }

    const padding = Math.max(1, Math.round(scale * 4))
    const left = Math.max(0, bounds.left - padding)
    const top = Math.max(0, bounds.top - padding)
    const right = Math.min(width, bounds.left + bounds.width + padding)
    const bottom = Math.min(height, bounds.top + bounds.height + padding)
    const croppedWidth = right - left
    const croppedHeight = bottom - top
    const logicalWidth = croppedWidth / scale
    const logicalHeight = croppedHeight / scale
    const dimensions = getRasterDimensions(logicalWidth, logicalHeight, sizeOrWidth, quality)
    canvas.width = 1
    canvas.height = 1
    const output = window.document.createElement('canvas')
    output.width = dimensions.width
    output.height = dimensions.height
    try {
      const outputContext = output.getContext('2d')
      if (!outputContext) {
        throw new Error('This browser cannot crop the PNG fallback.')
      }
      // Rasterize the vector source at the final resolution, not the bounds-probe
      // bitmap: enlarging that bitmap would add pixels without adding detail.
      outputContext.drawImage(
        image,
        left / scale,
        top / scale,
        logicalWidth,
        logicalHeight,
        0,
        0,
        output.width,
        output.height,
      )

      const dataUrl = output.toDataURL('image/png')
      const base64 = dataUrl.split(',', 2)[1]
      if (!dataUrl.startsWith('data:image/png;base64,') || !base64) {
        throw new Error('This browser cannot export the diagram at this PNG resolution.')
      }
      return {
        base64,
        width: logicalWidth,
        height: logicalHeight,
      }
    } finally {
      // Release the large backing store immediately between live updates.
      output.width = 1
      output.height = 1
    }
  } finally {
    URL.revokeObjectURL(url)
  }
}

export async function svgToPngBase64(svg: string, quality: ImageQuality = 'auto'): Promise<string> {
  return (await rasterizeSvg(svg, undefined, quality)).base64
}
