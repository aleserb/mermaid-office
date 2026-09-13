import type { DiagramSize } from '../metadata/payload'

// Logical display widths are in points, independently of raster pixel density.
export const DIAGRAM_WIDTHS: Record<DiagramSize, number> = {
  small: 216,
  medium: 324,
  large: 396,
  'page-width': 468,
}

export function fitDiagram(
  width: number,
  height: number,
  maxWidth = 500,
  maxHeight = 650,
  allowUpscale = false,
): { width: number; height: number } {
  const scale = Math.min(
    allowUpscale ? Number.POSITIVE_INFINITY : 1,
    maxWidth / width,
    maxHeight / height,
  )
  return {
    width: width * scale,
    height: height * scale,
  }
}
