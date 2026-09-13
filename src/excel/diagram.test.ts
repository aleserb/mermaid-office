import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import { readPayloadFromImage } from '../metadata/imageMetadata'
import { embedPayloadInPng } from '../metadata/pngMetadata'
import {
  createDiagramPayload,
  getDocumentSettingKey,
  getExcelShapeName,
} from '../metadata/payload'
import { rasterizeSvg } from '../word/insertDiagram'
import {
  getSelectedDiagram,
  insertDiagramWithPayload,
  updateDiagramById,
} from './diagram'

vi.mock('../metadata/imageMetadata', () => ({ readPayloadFromImage: vi.fn() }))
vi.mock('../metadata/pngMetadata', () => ({ embedPayloadInPng: vi.fn(() => 'embedded-png') }))
vi.mock('../word/insertDiagram', () => ({
  fitDiagram: vi.fn((width: number, height: number, maxWidth: number) => ({
    width: maxWidth,
    height: maxWidth * height / width,
  })),
  rasterizeSvg: vi.fn().mockResolvedValue({ base64: 'png', width: 600, height: 400 }),
}))

interface FakeShape {
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
const saveAsync = vi.fn((callback: (result: { status: string }) => void) => {
  callback({ status: 'succeeded' })
})
let activeShape: FakeShape
let shapes: FakeShape[]
let inserted: FakeShape[]

function shape(overrides: Partial<FakeShape> = {}): FakeShape {
  return {
    isNullObject: false,
    name: '',
    type: 'Image',
    left: 0,
    top: 0,
    width: 300,
    height: 200,
    load: vi.fn(),
    delete: vi.fn(),
    getAsImage: vi.fn(() => ({ value: 'shape-png' })),
    ...overrides,
  }
}

function setupExcel() {
  inserted = []
  const worksheet = {
    shapes: {
      addImage: vi.fn(() => {
        const result = shape()
        inserted.push(result)
        return result
      }),
      getItemOrNullObject: vi.fn((name: string) =>
        shapes.find((candidate) => candidate.name === name) ??
        shape({ isNullObject: true, name })),
    },
  }
  const worksheets = {
    items: [worksheet],
    load: vi.fn(),
    getActiveWorksheet: vi.fn(() => worksheet),
  }
  const workbook = {
    worksheets,
    getActiveShapeOrNullObject: vi.fn(() => activeShape),
    getSelectedRange: vi.fn(() => ({ left: 72, top: 144, load: vi.fn() })),
  }
  vi.stubGlobal('Excel', {
    PictureFormat: { png: 'PNG' },
    run: vi.fn(async (callback: (context: unknown) => unknown) =>
      callback({ workbook, sync: vi.fn().mockResolvedValue(undefined) })),
  })
  vi.stubGlobal('Office', {
    AsyncResultStatus: { Failed: 'failed' },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      document: {
        settings: {
          get: vi.fn((key: string) => values.get(key)),
          set: vi.fn((key: string, value: string) => values.set(key, value)),
          saveAsync,
        },
      },
    },
  })
  return { workbook, worksheet }
}

beforeEach(() => {
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
  vi.unstubAllGlobals()
  vi.clearAllMocks()
})

it('inserts a named PNG over the selected cell and stores its payload', async () => {
  const payload = await insertDiagramWithPayload(
    '<svg/>',
    'flowchart LR\nA-->B',
    'forest',
    'medium',
    { base64: 'png', width: 600, height: 400 },
    {},
    DEFAULT_DIAGRAM_SETTINGS,
  )

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

  await expect(updateDiagramById(
    '<svg/>',
    payload,
    'flowchart LR\nA-->C',
    'dark',
    'large',
    false,
    { base64: 'updated', width: 800, height: 400 },
    DEFAULT_DIAGRAM_SETTINGS,
  )).resolves.toBe('png')

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
  await expect(updateDiagramById('<svg/>', payload, payload.source)).rejects.toThrow(
    'deleted or can no longer be found',
  )

  values.set(getDocumentSettingKey(payload.id), JSON.stringify({ ...payload, source: 'changed' }))
  await expect(updateDiagramById('<svg/>', payload, payload.source)).rejects.toThrow(
    'changed in another editor',
  )
})
