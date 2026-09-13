import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { act, cleanup, renderHook } from '@testing-library/react'
import { LIVE_UPDATE_DELAY, usePaneEditor } from '../pane/usePaneEditor'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import { readPayloadFromImage } from '../metadata/imageMetadata'
import { embedPayloadInPng } from '../metadata/pngMetadata'
import {
  createDiagramPayload,
  getDocumentSettingKey,
  getExcelShapeName,
} from '../metadata/payload'
import { rasterizeSvg } from '../rendering/rasterizeSvg'
import type { UpdateDiagramRequest } from '../office/diagramRequests'
import {
  getSelectedDiagram,
  insertDiagramWithPayload,
  updateDiagramById,
} from './diagram'

vi.mock('../metadata/imageMetadata', () => ({ readPayloadFromImage: vi.fn() }))
vi.mock('../metadata/pngMetadata', () => ({ embedPayloadInPng: vi.fn(() => 'embedded-png') }))
vi.mock('../mermaid/render', async importOriginal => ({
  ...await importOriginal<typeof import('../mermaid/render')>(),
  renderMermaid: vi.fn().mockResolvedValue('<svg/>'),
}))
vi.mock('../word/insertDiagram', () => ({
  insertDiagramWithPayload: vi.fn(),
  updateDiagramById: vi.fn(),
}))
vi.mock('../rendering/rasterizeSvg', () => ({
  rasterizeSvg: vi.fn().mockResolvedValue({ base64: 'png', width: 600, height: 400 }),
}))

interface FakeShape {
  id: string
  isNullObject: boolean
  name: string
  type: string
  left: number
  top: number
  width: number
  height: number
  load: ReturnType<typeof vi.fn>
  delete: ReturnType<typeof vi.fn>
  getAsImage: ReturnType<typeof vi.fn>
}

const values = new Map<string, string>()
const saveAsync = vi.fn((callback: (result: { status: string; error?: { message: string } }) => void) => {
  callback({ status: 'succeeded' })
})
const sync = vi.fn<() => Promise<void>>()
let activeShape: FakeShape
let shapes: FakeShape[]
let inserted: FakeShape[]
let nextShapeId = 0

function shape(overrides: Partial<FakeShape> = {}): FakeShape {
  const result: FakeShape = {
    id: `shape-${nextShapeId++}`,
    isNullObject: false,
    name: '',
    type: 'Image',
    left: 0,
    top: 0,
    width: 300,
    height: 200,
    load: vi.fn(),
    delete: vi.fn(() => {
      shapes = shapes.filter(candidate => candidate !== result)
      if (activeShape === result) activeShape = shape({ isNullObject: true })
    }),
    getAsImage: vi.fn(() => ({ value: 'shape-png' })),
    ...overrides,
  }
  return result
}

function setupExcel() {
  inserted = []
  const worksheet = {
    id: 'worksheet-1',
    load: vi.fn(),
    shapes: {
      addImage: vi.fn(() => {
        const result = shape()
        inserted.push(result)
        shapes.push(result)
        return result
      }),
      getItemOrNullObject: vi.fn((name: string) =>
        shapes.find((candidate) => candidate.name === name || candidate.id === name) ??
        shape({ isNullObject: true, name })),
    },
  }
  const worksheets = {
    items: [worksheet],
    load: vi.fn(),
    getActiveWorksheet: vi.fn(() => worksheet),
    getItem: vi.fn(() => worksheet),
  }
  const workbook = {
    worksheets,
    getActiveShapeOrNullObject: vi.fn(() => activeShape),
    getSelectedRange: vi.fn(() => ({ left: 72, top: 144, load: vi.fn() })),
  }
  vi.stubGlobal('Excel', {
    PictureFormat: { png: 'PNG' },
    run: vi.fn(async (callback: (context: unknown) => unknown) =>
      callback({ workbook, sync })),
  })
  vi.stubGlobal('Office', {
    onReady: vi.fn().mockResolvedValue({ host: 'Excel' }),
    EventType: { DocumentSelectionChanged: 'selection-change' },
    AsyncResultStatus: { Failed: 'failed' },
    context: {
      host: 'Excel',
      platform: 'PC',
      requirements: { isSetSupported: vi.fn(() => true) },
      document: {
        addHandlerAsync: vi.fn((_event, _handler, callback) => callback({ status: 'succeeded' })),
        removeHandlerAsync: vi.fn(),
        settings: {
          get: vi.fn((key: string) => values.get(key)),
          set: vi.fn((key: string, value: string) => values.set(key, value)),
          remove: vi.fn((key: string) => values.delete(key)),
          saveAsync,
        },
      },
    },
  })
  return { workbook, worksheet }
}

beforeEach(() => {
  sync.mockReset().mockResolvedValue(undefined)
  saveAsync.mockReset().mockImplementation(callback => callback({ status: 'succeeded' }))
  values.clear()
  shapes = []
  activeShape = shape({ isNullObject: true })
  setupExcel()
  vi.stubGlobal('crypto', { randomUUID: vi.fn(() => 'new-diagram') })
  vi.mocked(readPayloadFromImage).mockReset()
  vi.mocked(embedPayloadInPng).mockClear()
  vi.mocked(rasterizeSvg).mockClear()
})

afterEach(() => {
  cleanup()
  vi.useRealTimers()
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
  vi.clearAllMocks()
  localStorage.clear()
})

it('loads and edits an existing Excel image even when WebView focus stays true', async () => {
  vi.useFakeTimers()
  vi.spyOn(document, 'hasFocus').mockReturnValue(true)
  vi.spyOn(document, 'visibilityState', 'get').mockReturnValue('visible')
  const payload = createDiagramPayload('flowchart LR\nExisting-->Diagram', 'png')
  values.set(getDocumentSettingKey(payload.id), JSON.stringify(payload))
  const original = shape({ name: getExcelShapeName(payload.id) })
  shapes = [original]
  const { result } = renderHook(usePaneEditor)
  await act(async () => { await vi.advanceTimersByTimeAsync(0) })
  expect(result.current.target).toBeNull()

  // Selecting an image need not change the selected cells or emit an Office event.
  activeShape = original
  await act(async () => { await vi.advanceTimersByTimeAsync(500) })
  expect(result.current.target?.id).toBe(payload.id)
  expect(result.current.draft.source).toBe(payload.source)
  expect(result.current.loadingSelection).toBe(false)

  act(() => result.current.changeSource('flowchart LR\nExisting-->Updated'))
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
  expect(original.delete).toHaveBeenCalledOnce()
  expect(inserted).toHaveLength(1)
  expect(JSON.parse(values.get(getDocumentSettingKey(payload.id))!)).toMatchObject({
    source: 'flowchart LR\nExisting-->Updated',
  })
  expect(result.current.target?.source).toBe('flowchart LR\nExisting-->Updated')
})

it('reloads updated source after inserting, deselecting, and reselecting an image', async () => {
  vi.useFakeTimers()
  vi.spyOn(document, 'hasFocus').mockReturnValue(true)
  vi.spyOn(document, 'visibilityState', 'get').mockReturnValue('visible')
  const { result } = renderHook(usePaneEditor)
  await act(async () => { await vi.advanceTimersByTimeAsync(0) })
  act(() => result.current.changeSource('flowchart LR\nNew-->Diagram'))
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
  await act(async () => { await result.current.insert() })
  const payload = result.current.target!
  activeShape = inserted[0]
  shapes = [activeShape]
  await act(async () => { await vi.advanceTimersByTimeAsync(500) })
  act(() => result.current.changeSource('flowchart LR\nNew-->Updated'))
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
  const replacement = inserted[1]
  shapes = [replacement]
  activeShape = shape({ isNullObject: true })
  vi.mocked(document.hasFocus).mockReturnValue(false)
  await act(async () => {
    await vi.advanceTimersByTimeAsync(500)
  })
  expect(result.current.target).toBeNull()

  // A shape click changes neither the cell range nor the reported WebView focus.
  vi.mocked(document.hasFocus).mockReturnValue(true)
  activeShape = replacement
  await act(async () => { await vi.advanceTimersByTimeAsync(500) })
  expect(result.current.target?.id).toBe(payload.id)
  expect(result.current.draft.source).toBe('flowchart LR\nNew-->Updated')
  expect(result.current.loadingSelection).toBe(false)
  act(() => result.current.changeSource('flowchart LR\nReselected-->Edited'))
  await act(async () => { await vi.advanceTimersByTimeAsync(LIVE_UPDATE_DELAY) })
  expect(replacement.delete).toHaveBeenCalledOnce()
})

it('inserts a named PNG over the selected cell and stores its payload', async () => {
  const payload = await insertDiagramWithPayload({
    svg: '<svg/>',
    draft: {
      source: 'flowchart LR\nA-->B', theme: 'forest', size: 'medium', settings: DEFAULT_DIAGRAM_SETTINGS,
    },
    raster: { base64: 'png', width: 600, height: 400 },
  })

  expect(payload.id).toBe('new-diagram')
  expect(inserted).toHaveLength(1)
  expect(inserted[0]).toMatchObject({
    name: getExcelShapeName(payload.id),
    left: 72,
    top: 144,
    width: 324,
    height: 216,
  })
  expect(embedPayloadInPng).toHaveBeenCalledWith('png', payload)
  expect(JSON.parse(values.get(getDocumentSettingKey(payload.id))!)).toEqual(payload)
  expect(saveAsync).toHaveBeenCalledOnce()
})

it('loads a selected managed image from workbook settings', async () => {
  const payload = createDiagramPayload('flowchart LR\nA-->B', 'png')
  activeShape = shape({ name: getExcelShapeName(payload.id) })
  shapes = [activeShape]
  values.set(getDocumentSettingKey(payload.id), JSON.stringify(payload))

  await expect(getSelectedDiagram()).resolves.toEqual(payload)
  expect(activeShape.getAsImage).not.toHaveBeenCalled()
})

it('recovers embedded metadata and assigns copied shapes a new identity', async () => {
  const payload = createDiagramPayload('flowchart LR\nA-->B', 'png')
  activeShape = shape({ name: getExcelShapeName(payload.id) })
  shapes = [activeShape, shape({ name: getExcelShapeName(payload.id) })]
  vi.mocked(readPayloadFromImage).mockReturnValue(payload)

  await expect(getSelectedDiagram()).resolves.toMatchObject({ id: 'new-diagram' })
  expect(activeShape.name).toBe(getExcelShapeName('new-diagram'))
  expect(readPayloadFromImage).toHaveBeenCalledWith('shape-png')
  expect(values.has(getDocumentSettingKey('new-diagram'))).toBe(true)
})

it('replaces the stored shape while preserving position and displayed width', async () => {
  const payload = createDiagramPayload('flowchart LR\nA-->B', 'png')
  const original = shape({
    name: getExcelShapeName(payload.id),
    left: 18,
    top: 36,
    width: 360,
    height: 240,
  })
  shapes = [original]
  values.set(getDocumentSettingKey(payload.id), JSON.stringify(payload))

  await expect(updateDiagramById({
    svg: '<svg/>',
    existing: payload,
    draft: {
      source: 'flowchart LR\nA-->C', theme: 'dark', size: 'large', settings: DEFAULT_DIAGRAM_SETTINGS,
    },
    applySize: false,
    raster: { base64: 'updated', width: 800, height: 400 },
  })).resolves.toBe('png')

  expect(original.delete).toHaveBeenCalledOnce()
  expect(inserted[0]).toMatchObject({
    name: getExcelShapeName(payload.id),
    left: 18,
    top: 36,
    width: 360,
    height: 180,
  })
  const stored = JSON.parse(values.get(getDocumentSettingKey(payload.id))!)
  expect(stored).toMatchObject({ source: 'flowchart LR\nA-->C', theme: 'dark', size: 'large' })
})

it('rejects missing and concurrently changed diagrams', async () => {
  const payload = createDiagramPayload('flowchart LR\nA-->B', 'png')
  const request: UpdateDiagramRequest = {
    svg: '<svg/>', existing: payload, draft: { ...payload, settings: DEFAULT_DIAGRAM_SETTINGS },
  }
  await expect(updateDiagramById(request)).rejects.toThrow(
    'deleted or can no longer be found',
  )

  values.set(getDocumentSettingKey(payload.id), JSON.stringify({ ...payload, source: 'changed' }))
  await expect(updateDiagramById(request)).rejects.toThrow(
    'changed in another editor',
  )
})

function replacementFixture() {
  const payload = createDiagramPayload('flowchart LR\nOriginal-->Diagram', 'png')
  const original = shape({ name: getExcelShapeName(payload.id) })
  shapes = [original]
  const serialized = JSON.stringify(payload)
  values.set(getDocumentSettingKey(payload.id), serialized)
  const request: UpdateDiagramRequest = {
    svg: '<svg/>', existing: payload,
    draft: { ...payload, source: 'flowchart LR\nUpdated-->Diagram', settings: DEFAULT_DIAGRAM_SETTINGS },
    raster: { base64: 'png', width: 600, height: 400 },
  }
  return { payload, original, serialized, request }
}

it('retains the original image until replacement insertion and metadata saving succeed', async () => {
  const { original, request } = replacementFixture()
  saveAsync.mockImplementationOnce(callback => {
    expect(original.delete).not.toHaveBeenCalled()
    expect(shapes).toContain(original)
    expect(inserted).toHaveLength(1)
    expect(inserted[0].name).toMatch(/^Mermaid pending /)
    callback({ status: 'succeeded' })
  })
  await updateDiagramById(request)
  expect(shapes).toEqual([inserted[0]])
})

it('does not mutate the workbook when image preparation fails', async () => {
  const { original, request } = replacementFixture()
  vi.mocked(rasterizeSvg).mockRejectedValueOnce(new Error('Image preparation failed'))
  await expect(updateDiagramById({ ...request, raster: undefined })).rejects.toThrow('Image preparation failed')
  expect(shapes).toEqual([original])
  expect(inserted).toHaveLength(0)
  expect(saveAsync).not.toHaveBeenCalled()
})

it('does not overwrite an external edit made during image preparation', async () => {
  const { original, request, payload } = replacementFixture()
  const external = JSON.stringify({ ...payload, source: 'flowchart LR\nExternal-->Edit' })
  vi.mocked(embedPayloadInPng).mockImplementationOnce(() => {
    values.set(getDocumentSettingKey(payload.id), external)
    return 'embedded-png'
  })
  await expect(updateDiagramById(request)).rejects.toThrow('changed in another editor')
  expect(shapes).toEqual([original])
  expect(values.get(getDocumentSettingKey(payload.id))).toBe(external)
  expect(saveAsync).not.toHaveBeenCalled()
})

it('keeps the original and old metadata when replacement insertion fails', async () => {
  const { original, serialized, request, payload } = replacementFixture()
  // Workbook lookup, shape lookup, worksheet ID, then staged image creation.
  sync.mockResolvedValueOnce().mockResolvedValueOnce().mockResolvedValueOnce()
    .mockRejectedValueOnce(new Error('Image insertion failed'))
  await expect(updateDiagramById(request)).rejects.toThrow('Image insertion failed')
  expect(original.delete).not.toHaveBeenCalled()
  expect(shapes).toEqual([original])
  expect(values.get(getDocumentSettingKey(payload.id))).toBe(serialized)
  expect(saveAsync).not.toHaveBeenCalled()
})

it('removes the staged image and restores metadata when saving fails', async () => {
  const { original, serialized, request, payload } = replacementFixture()
  saveAsync.mockImplementationOnce(callback =>
    callback({ status: 'failed', error: { message: 'Settings save failed' } }))
  await expect(updateDiagramById(request)).rejects.toThrow('Settings save failed')
  expect(original.delete).not.toHaveBeenCalled()
  expect(shapes).toEqual([original])
  expect(original.name).toBe(getExcelShapeName(payload.id))
  expect(values.get(getDocumentSettingKey(payload.id))).toBe(serialized)
  expect(saveAsync).toHaveBeenCalledTimes(2)
})

it('restores an absent settings entry when an update of embedded-only metadata fails', async () => {
  const { original, request, payload } = replacementFixture()
  values.delete(getDocumentSettingKey(payload.id))
  saveAsync.mockImplementationOnce(callback =>
    callback({ status: 'failed', error: { message: 'Settings save failed' } }))
  await expect(updateDiagramById(request)).rejects.toThrow('Settings save failed')
  expect(shapes).toEqual([original])
  expect(values.has(getDocumentSettingKey(payload.id))).toBe(false)
})

it('restores the original name and source after a partially applied name swap', async () => {
  const { original, serialized, request, payload } = replacementFixture()
  sync.mockResolvedValueOnce().mockResolvedValueOnce().mockResolvedValueOnce()
    .mockResolvedValueOnce().mockResolvedValueOnce()
    .mockRejectedValueOnce(new Error('Name swap failed'))
  await expect(updateDiagramById(request)).rejects.toThrow('Name swap failed')
  expect(shapes).toEqual([original])
  expect(original.name).toBe(getExcelShapeName(payload.id))
  expect(values.get(getDocumentSettingKey(payload.id))).toBe(serialized)
})

it('rolls back when deletion fails before removing the original', async () => {
  const { original, serialized, request, payload } = replacementFixture()
  original.delete.mockImplementationOnce(() => { throw new Error('Delete failed') })
  await expect(updateDiagramById(request)).rejects.toThrow('Delete failed')
  expect(shapes).toEqual([original])
  expect(original.name).toBe(getExcelShapeName(payload.id))
  expect(values.get(getDocumentSettingKey(payload.id))).toBe(serialized)
})

it('confirms committed state if the final sync rejects after deleting the original', async () => {
  const { request, payload } = replacementFixture()
  const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
  sync.mockResolvedValueOnce().mockResolvedValueOnce().mockResolvedValueOnce()
    .mockResolvedValueOnce().mockResolvedValueOnce().mockResolvedValueOnce()
    .mockRejectedValueOnce(new Error('Final sync acknowledgement lost'))
  await expect(updateDiagramById(request)).resolves.toBe('png')
  expect(shapes).toEqual([inserted[0]])
  expect(inserted[0].name).toBe(getExcelShapeName(payload.id))
  expect(JSON.parse(values.get(getDocumentSettingKey(payload.id))!).source).toBe(request.draft.source)
  expect(warning).toHaveBeenCalledOnce()
})

it('surfaces incomplete recovery without deleting the original image', async () => {
  const { original, request } = replacementFixture()
  saveAsync.mockImplementation(callback =>
    callback({ status: 'failed', error: { message: 'Settings unavailable' } }))
  await expect(updateDiagramById(request)).rejects.toThrow('fully recover')
  expect(shapes).toEqual([original])
  expect(original.delete).not.toHaveBeenCalled()
})
