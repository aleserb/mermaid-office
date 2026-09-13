import { getSvgDimensions } from '../mermaid/svgDimensions'

export function requiresManualLargeDiagramUpdates(platform?: string): boolean {
  return platform !== 'PC' && platform !== 'Mac'
}

export function isLargeDiagram(source: string, svg?: string): boolean {
  if (source.length >= 4000 || source.split(/\r?\n/).filter(line => line.trim()).length >= 50) {
    return true
  }
  if (!svg) return false
  const root = new DOMParser().parseFromString(svg, 'image/svg+xml').documentElement
  const { width, height } = getSvgDimensions(root)
  return width > 1536 || height > 1536 || width * height > 2 * 1024 * 1024
}
