import { afterEach, expect, it, vi } from 'vitest'
import * as excel from '../excel/diagram'
import { DEFAULT_DIAGRAM_SETTINGS } from '../metadata/diagramSettings'
import { createDiagramPayload } from '../metadata/payload'
import * as word from '../word/insertDiagram'
import type { InsertDiagramRequest, UpdateDiagramRequest } from './diagramRequests'
import { getHostAdapter } from './hostAdapter'

vi.mock('../word/insertDiagram', () => ({
  insertDiagramWithPayload: vi.fn(),
  updateDiagramById: vi.fn(),
}))
vi.mock('../excel/diagram', () => ({
  getSelectedDiagram: vi.fn(),
  watchSelectedDiagram: vi.fn(),
  insertDiagramWithPayload: vi.fn(),
  updateDiagramById: vi.fn(),
}))

afterEach(() => vi.clearAllMocks())

it('routes supported Office hosts to their document adapter', () => {
  expect(getHostAdapter().appName).toBe('Word')
  expect(getHostAdapter('Word').appName).toBe('Word')
  expect(getHostAdapter('Excel').appName).toBe('Excel')
})

it('rejects unsupported Office hosts', () => {
  expect(() => getHostAdapter('PowerPoint')).toThrow(
    'Mermaid Office supports Microsoft Word and Excel.',
  )
})

it.each([
  { host: 'Word', implementation: word },
  { host: 'Excel', implementation: excel },
])('forwards one typed request unchanged to $host', async ({ host, implementation }) => {
  const payload = createDiagramPayload('flowchart LR\nA --> B', 'png')
  const raster = { base64: 'prepared-png', width: 100, height: 50 }
  const draft = {
    source: payload.source, theme: payload.theme, size: payload.size,
    settings: DEFAULT_DIAGRAM_SETTINGS,
  }
  const insert: InsertDiagramRequest = {
    svg: '<svg/>', draft, raster, requireEmptySelection: true,
  }
  const update: UpdateDiagramRequest = {
    svg: '<svg/>', existing: payload, draft, raster, applySize: true,
  }
  vi.mocked(implementation.insertDiagramWithPayload).mockResolvedValueOnce(payload)
  vi.mocked(implementation.updateDiagramById).mockResolvedValueOnce('png')

  const adapter = getHostAdapter(host)
  await expect(adapter.insertDiagramWithPayload(insert)).resolves.toBe(payload)
  await expect(adapter.updateDiagramById(update)).resolves.toBe('png')
  expect(implementation.insertDiagramWithPayload).toHaveBeenCalledExactlyOnceWith(insert)
  expect(implementation.updateDiagramById).toHaveBeenCalledExactlyOnceWith(update)
})
