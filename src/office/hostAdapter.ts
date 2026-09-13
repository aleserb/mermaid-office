import type { DiagramSettings } from '../metadata/diagramSettings'
import type {
  DiagramFormat,
  DiagramPayload,
  DiagramSize,
  DiagramTheme,
} from '../metadata/payload'
import type {
  DiagramInsertionOptions,
  RasterizedDiagram,
} from '../word/insertDiagram'
import type { DiagramSelectionWatcher, SelectionWatchOptions } from './selectionWatcher'
import * as excel from '../excel/diagram'
import * as wordInsert from '../word/insertDiagram'
import * as wordSelection from '../word/selection'

export interface HostAdapter {
  appName: 'Word' | 'Excel'
  insertionLocationError: string
  getSelectedDiagram(): Promise<DiagramPayload | null>
  watchSelectedDiagram(
    onSelected: (payload: DiagramPayload | null) => void,
    onError: (error: Error) => void,
    options?: SelectionWatchOptions,
  ): DiagramSelectionWatcher
  insertDiagramWithPayload(
    svg: string,
    source: string,
    theme?: DiagramTheme,
    size?: DiagramSize,
    renderedRaster?: RasterizedDiagram,
    options?: DiagramInsertionOptions,
    settings?: DiagramSettings,
  ): Promise<DiagramPayload>
  updateDiagramById(
    svg: string,
    existing: DiagramPayload,
    source: string,
    theme?: DiagramTheme,
    size?: DiagramSize,
    applySize?: boolean,
    renderedRaster?: RasterizedDiagram,
    settings?: DiagramSettings,
  ): Promise<DiagramFormat>
}

const wordAdapter: HostAdapter = {
  appName: 'Word',
  insertionLocationError: 'Place the cursor on a blank line before inserting a new diagram.',
  getSelectedDiagram: wordSelection.getSelectedDiagram,
  watchSelectedDiagram: wordSelection.watchSelectedDiagram,
  insertDiagramWithPayload: wordInsert.insertDiagramWithPayload,
  updateDiagramById: wordInsert.updateDiagramById,
}

const excelAdapter: HostAdapter = {
  appName: 'Excel',
  insertionLocationError: 'Select a cell before inserting a new diagram.',
  ...excel,
}

export function getHostAdapter(host?: Office.HostType | string): HostAdapter {
  const name = host === undefined ? 'Word' : String(host)
  if (name === 'Word') return wordAdapter
  if (name === 'Excel') return excelAdapter
  throw new Error('Mermaid Office supports Microsoft Word and Excel.')
}
