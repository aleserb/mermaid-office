import type { DiagramSettings } from '../metadata/diagramSettings'
import type { DiagramPayload, DiagramSize, DiagramTheme } from '../metadata/payload'
import type { RasterizedDiagram } from '../rendering/rasterizeSvg'

export interface DiagramDraft {
  source: string
  theme: DiagramTheme
  size: DiagramSize
  settings: DiagramSettings
}

export interface InsertDiagramRequest {
  svg: string
  draft: DiagramDraft
  raster?: RasterizedDiagram
  requireEmptySelection?: boolean
}

export interface UpdateDiagramRequest {
  svg: string
  existing: DiagramPayload
  draft: DiagramDraft
  applySize?: boolean
  raster?: RasterizedDiagram
}
