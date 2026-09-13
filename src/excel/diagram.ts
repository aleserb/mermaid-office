import { getDiagramSettings } from '../metadata/diagramSettings'
import { readPayloadFromImage } from '../metadata/imageMetadata'
import { embedPayloadInPng } from '../metadata/pngMetadata'
import {
  createDiagramPayload,
  getDiagramIdFromExcelShapeName,
  getDocumentSettingKey,
  getExcelShapeName,
  parseDiagramPayload,
  sameDiagramPayload,
  type DiagramFormat,
  type DiagramPayload,
} from '../metadata/payload'
import { watchDiagramSelection, type DiagramSelectionWatcher, type SelectionReceiver, type SelectionWatchOptions } from '../office/selectionWatcher'
import type { InsertDiagramRequest, UpdateDiagramRequest } from '../office/diagramRequests'
import { DIAGRAM_WIDTHS, fitDiagram } from '../rendering/diagramSizing'
import { rasterizeSvg } from '../rendering/rasterizeSvg'

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
  await flushSettings()
}

async function flushSettings(): Promise<void> {
  const settings = getSettings()
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

async function findShapesByName(
  context: Excel.RequestContext,
  name: string,
): Promise<Array<{ shape: Excel.Shape; worksheet: Excel.Worksheet }>> {
  const worksheets = context.workbook.worksheets
  worksheets.load('items')
  await context.sync()
  const candidates = worksheets.items.map((worksheet) => {
    const shape = worksheet.shapes.getItemOrNullObject(name)
    shape.load('isNullObject,id,left,top,width,height,name')
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
  onSelected: SelectionReceiver,
  onError: (error: Error) => void,
  options: SelectionWatchOptions = {},
): DiagramSelectionWatcher {
  return watchDiagramSelection(getSelectedDiagram, onSelected, onError, {
    ...options,
    pollIntervalMs: 500,
  })
}

export async function insertDiagramWithPayload({
  svg, draft: { source, theme, size, settings }, raster: renderedRaster,
}: InsertDiagramRequest): Promise<DiagramPayload> {
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

export async function updateDiagramById({
  svg, existing, draft: { source, theme, size, settings }, applySize = false,
  raster: renderedRaster,
}: UpdateDiagramRequest): Promise<DiagramFormat> {
  ensureExcelApi()
  const stored = readStoredPayload(existing.id)
  if (stored && !sameDiagramPayload(stored, existing)) {
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
    await replaceShape(context, {
      shape, worksheet, png, dimensions, payload, existing,
    })
  })
  return 'png'
}

interface ReplacementState {
  worksheetId: string
  originalId: string
  replacementId?: string
  stagingName: string
  stableName: string
  settingKey: string
  previousSetting: unknown
  settingsTouched: boolean
  payloadSaved: boolean
}

async function replaceShape(context: Excel.RequestContext, {
  shape, worksheet, png, dimensions, payload, existing,
}: {
  shape: Excel.Shape
  worksheet: Excel.Worksheet
  png: string
  dimensions: { width: number; height: number }
  payload: DiagramPayload
  existing: DiagramPayload
}): Promise<void> {
  worksheet.load('id')
  await context.sync()
  const settings = getSettings()
  const settingKey = getDocumentSettingKey(payload.id)
  const state: ReplacementState = {
    worksheetId: worksheet.id,
    originalId: shape.id,
    stagingName: `Mermaid pending ${crypto.randomUUID()}`,
    stableName: getExcelShapeName(payload.id),
    settingKey,
    previousSetting: settings.get(settingKey),
    settingsTouched: false,
    payloadSaved: false,
  }
  try {
    // Keep the original intact until the new image AND its source are durable.
    const replacement = worksheet.shapes.addImage(png)
    replacement.name = state.stagingName
    replacement.left = shape.left
    replacement.top = shape.top
    replacement.width = dimensions.width
    replacement.height = dimensions.height
    replacement.load('id')
    await context.sync()
    state.replacementId = replacement.id

    const current = readStoredPayload(existing.id)
    if (current && !sameDiagramPayload(current, existing)) {
      throw new Error('The Mermaid diagram changed in another editor. Select it again before updating.')
    }
    state.settingsTouched = true
    await savePayload(payload)
    state.payloadSaved = true
    shape.name = `Mermaid previous ${crypto.randomUUID()}`
    await context.sync()
    replacement.name = state.stableName
    await context.sync()
    shape.delete()
    await context.sync()
  } catch (error) {
    let committed: boolean
    try {
      committed = await recoverReplacement(state)
    } catch (recoveryError) {
      throw new AggregateError([error, recoveryError],
        'Excel could not finish or fully recover the diagram update. Keep the workbook open and inspect the original and replacement images before retrying.')
    }
    if (!committed) throw error
    // A failed sync can still have applied its final delete. Confirm host state
    // before reporting success instead of deleting the only surviving image.
    console.warn('Excel reported an update error, but the replacement and saved source were confirmed.', error)
  }
}

async function recoverReplacement(state: ReplacementState): Promise<boolean> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(state.worksheetId)
    const original = worksheet.shapes.getItemOrNullObject(state.originalId)
    const replacement = worksheet.shapes.getItemOrNullObject(state.replacementId ?? state.stagingName)
    original.load('isNullObject')
    replacement.load('isNullObject,name')
    await context.sync()
    if (original.isNullObject) {
      if (!replacement.isNullObject && replacement.name === state.stableName && state.payloadSaved) {
        return true
      }
      throw new Error('The original diagram is unavailable and the replacement could not be confirmed.')
    }

    // Restore the name first, after freeing it if a partially applied swap took it.
    if (!replacement.isNullObject) {
      replacement.delete()
      await context.sync()
    }
    original.name = state.stableName
    await context.sync()
    if (state.settingsTouched) {
      const settings = getSettings()
      if (state.previousSetting === undefined || state.previousSetting === null) {
        settings.remove(state.settingKey)
      } else {
        settings.set(state.settingKey, state.previousSetting)
      }
      await flushSettings()
    }
    return false
  })
}
