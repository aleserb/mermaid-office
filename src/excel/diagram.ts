import { getDiagramSettings, sameDiagramSettings, type DiagramSettings } from '../metadata/diagramSettings'
import { readPayloadFromImage } from '../metadata/imageMetadata'
import { embedPayloadInPng } from '../metadata/pngMetadata'
import {
  createDiagramPayload,
  getDiagramIdFromExcelShapeName,
  getDocumentSettingKey,
  getExcelShapeName,
  parseDiagramPayload,
  type DiagramFormat,
  type DiagramPayload,
  type DiagramSize,
  type DiagramTheme,
} from '../metadata/payload'
import { watchDiagramSelection, type DiagramSelectionWatcher, type SelectionWatchOptions } from '../office/selectionWatcher'
import {
  fitDiagram,
  rasterizeSvg,
  type DiagramInsertionOptions,
  type RasterizedDiagram,
} from '../word/insertDiagram'

const DIAGRAM_WIDTHS: Record<DiagramSize, number> = {
  small: 216,
  medium: 324,
  large: 396,
  'page-width': 468,
}

function ensureExcelApi(): void {
  if (typeof Excel === 'undefined') {
    throw new Error('Open Mermaid Office inside Microsoft Excel to work with diagrams.')
  }
  if (
    Office.context.requirements?.isSetSupported &&
    !Office.context.requirements.isSetSupported('ExcelApi', '1.19')
  ) {
    throw new Error('Mermaid Office requires a Microsoft 365 version of Excel that supports ExcelApi 1.19.')
  }
}

function getSettings(): Office.Settings {
  const settings = Office.context.document.settings
  if (!settings) {
    throw new Error('Excel workbook settings are unavailable.')
  }
  return settings
}

function readStoredPayload(id: string): DiagramPayload | null {
  const value = getSettings().get(getDocumentSettingKey(id))
  return value === undefined || value === null ? null : parseDiagramPayload(String(value))
}

async function savePayload(payload: DiagramPayload): Promise<void> {
  const settings = getSettings()
  settings.set(getDocumentSettingKey(payload.id), JSON.stringify(payload))
  await new Promise<void>((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message))
      } else {
        resolve()
      }
    })
  })
}

function samePayload(left: DiagramPayload, right: DiagramPayload): boolean {
  return left.id === right.id &&
    left.source === right.source &&
    left.theme === right.theme &&
    left.size === right.size &&
    left.format === right.format &&
    sameDiagramSettings(left.settings, right.settings)
}

async function findShapesByName(
  context: Excel.RequestContext,
  name: string,
): Promise<Array<{ shape: Excel.Shape; worksheet: Excel.Worksheet }>> {
  const worksheets = context.workbook.worksheets
  worksheets.load('items')
  await context.sync()
  const candidates = worksheets.items.map((worksheet) => {
    const shape = worksheet.shapes.getItemOrNullObject(name)
    shape.load('isNullObject,left,top,width,height,name')
    return { shape, worksheet }
  })
  await context.sync()
  return candidates.filter(({ shape }) => !shape.isNullObject)
}

export async function getSelectedDiagram(): Promise<DiagramPayload | null> {
  ensureExcelApi()
  const result = await Excel.run(async (context) => {
    const shape = context.workbook.getActiveShapeOrNullObject()
    shape.load('isNullObject,name,type')
    await context.sync()
    if (shape.isNullObject || String(shape.type).toLowerCase() !== 'image') {
      return null
    }

    const id = getDiagramIdFromExcelShapeName(shape.name)
    if (!id) {
      return null
    }

    const stored = readStoredPayload(id)
    const matches = await findShapesByName(context, shape.name)
    if (stored && matches.length === 1) {
      return { payload: stored, recovered: false }
    }

    let recovered = stored
    if (!recovered) {
      const image = shape.getAsImage(Excel.PictureFormat.png)
      await context.sync()
      recovered = readPayloadFromImage(image.value)
    }
    if (!recovered) {
      return null
    }

    const payload = matches.length > 1
      ? { ...recovered, id: crypto.randomUUID() }
      : { ...recovered, id }
    shape.name = getExcelShapeName(payload.id)
    await context.sync()
    return { payload, recovered: true }
  })

  if (!result) {
    return null
  }
  if (result.recovered) {
    await savePayload(result.payload)
  }
  return result.payload
}

export function watchSelectedDiagram(
  onSelected: (payload: DiagramPayload | null) => void,
  onError: (error: Error) => void,
  options: SelectionWatchOptions = {},
): DiagramSelectionWatcher {
  return watchDiagramSelection(getSelectedDiagram, onSelected, onError, {
    ...options,
    pollIntervalMs: 500,
  })
}

export async function insertDiagramWithPayload(
  svg: string,
  source: string,
  theme: DiagramTheme = 'default',
  size: DiagramSize = 'medium',
  renderedRaster?: RasterizedDiagram,
  _options: DiagramInsertionOptions = {},
  settings?: DiagramSettings,
): Promise<DiagramPayload> {
  ensureExcelApi()
  const payload = createDiagramPayload(source, 'png', theme, size, settings)
  const raster = renderedRaster ??
    await rasterizeSvg(svg, size, getDiagramSettings(settings).imageQuality)
  const png = embedPayloadInPng(raster.base64, payload)
  const dimensions = fitDiagram(raster.width, raster.height, DIAGRAM_WIDTHS[size], 650, true)

  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet()
    const range = context.workbook.getSelectedRange()
    range.load('left,top')
    await context.sync()
    const shape = worksheet.shapes.addImage(png)
    shape.name = getExcelShapeName(payload.id)
    shape.left = range.left
    shape.top = range.top
    shape.width = dimensions.width
    shape.height = dimensions.height
    await context.sync()
  })
  await savePayload(payload)
  return payload
}

export async function updateDiagramById(
  svg: string,
  existing: DiagramPayload,
  source: string,
  theme: DiagramTheme = existing.theme,
  size: DiagramSize = existing.size,
  applySize = false,
  renderedRaster?: RasterizedDiagram,
  settings?: DiagramSettings,
): Promise<DiagramFormat> {
  ensureExcelApi()
  const stored = readStoredPayload(existing.id)
  if (stored && !samePayload(stored, existing)) {
    throw new Error('The Mermaid diagram changed in another editor. Select it again before updating.')
  }

  const payload: DiagramPayload = {
    ...existing,
    source,
    theme,
    size,
    format: 'png',
    ...(settings ? { settings: { ...settings } } : {}),
  }

  await Excel.run(async (context) => {
    const matches = await findShapesByName(context, getExcelShapeName(existing.id))
    if (matches.length === 0) {
      throw new Error('The Mermaid diagram was deleted or can no longer be found.')
    }
    if (matches.length !== 1) {
      throw new Error('Multiple Mermaid diagrams share this ID. Select the diagram again to edit it.')
    }

    const { shape, worksheet } = matches[0]
    const raster = renderedRaster ?? await rasterizeSvg(
      svg,
      applySize ? size : shape.width,
      getDiagramSettings(payload.settings).imageQuality,
    )
    const dimensions = applySize
      ? fitDiagram(raster.width, raster.height, DIAGRAM_WIDTHS[size], 650, true)
      : { width: shape.width, height: shape.width * (raster.height / raster.width) }
    const png = embedPayloadInPng(raster.base64, payload)
    const position = { left: shape.left, top: shape.top }

    shape.delete()
    await context.sync()
    const replacement = worksheet.shapes.addImage(png)
    replacement.name = getExcelShapeName(payload.id)
    replacement.left = position.left
    replacement.top = position.top
    replacement.width = dimensions.width
    replacement.height = dimensions.height
    await context.sync()
  })
  await savePayload(payload)
  return 'png'
}
