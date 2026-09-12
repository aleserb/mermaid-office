import { afterEach, expect, it, vi } from 'vitest'
import { createDiagramPayload } from '../metadata/payload'
import { watchSelectedDiagram } from './selection'

afterEach(() => vi.unstubAllGlobals())

function setup() {
  let change: () => void = () => { throw new Error('Selection listener was not registered') }
  const read = vi.fn().mockResolvedValue(null)
  const remove = vi.fn()
  vi.stubGlobal('Word', { run: read })
  vi.stubGlobal('Office', {
    context: {
      document: {
        addHandlerAsync: (_event: string, handler: () => void, callback: (result: { status: string }) => void) => {
          change = handler
          callback({ status: 'succeeded' })
        },
        removeHandlerAsync: remove,
      },
    },
    EventType: { DocumentSelectionChanged: 'selection-change' },
    AsyncResultStatus: { Failed: 'failed' },
  })
  return { read, remove, change: () => change() }
}

it('drops stale lookups when another selection change arrives', async () => {
  const mock = setup()
  const oldDiagram = createDiagramPayload('flowchart LR\nOld', 'png')
  const newDiagram = createDiagramPayload('flowchart LR\nNew', 'png')
  let finish: (value: typeof oldDiagram) => void = () => { throw new Error('Lookup has not started') }
  mock.read.mockImplementationOnce(() => new Promise((resolve) => { finish = resolve }))
    .mockResolvedValueOnce(newDiagram)
  const selected = vi.fn()
  const changed = vi.fn()
  const stop = watchSelectedDiagram(selected, vi.fn(), { onSelectionChange: changed })
  mock.change()
  finish(oldDiagram)
  await vi.waitFor(() => expect(selected).toHaveBeenCalledOnce())
  expect(selected).toHaveBeenCalledWith(newDiagram)
  expect(changed).toHaveBeenCalledOnce()
  stop()
})

it('defers lookups during document writes and can refresh afterwards', async () => {
  const mock = setup()
  let paused = true
  const selected = vi.fn()
  const stop = watchSelectedDiagram(selected, vi.fn(), { isPaused: () => paused })
  mock.change()
  expect(mock.read).not.toHaveBeenCalled()
  paused = false
  stop.refresh()
  await vi.waitFor(() => expect(selected).toHaveBeenCalledWith(null))
  expect(mock.read).toHaveBeenCalledOnce()
  stop()
})

it('does not deliver a lookup that finishes after a document write starts', async () => {
  const mock = setup()
  let paused = false
  let finish: (value: null) => void = () => { throw new Error('Lookup has not started') }
  mock.read.mockImplementationOnce(() => new Promise((resolve) => { finish = resolve }))
  const selected = vi.fn()
  const stop = watchSelectedDiagram(selected, vi.fn(), { isPaused: () => paused })
  paused = true
  finish(null)
  await new Promise((resolve) => setTimeout(resolve, 0))
  expect(selected).not.toHaveBeenCalled()
  paused = false
  stop.refresh()
  await vi.waitFor(() => expect(selected).toHaveBeenCalledOnce())
  stop()
})

it('removes the handler and ignores late results after stopping', async () => {
  const mock = setup()
  let finish: (value: null) => void = () => { throw new Error('Lookup has not started') }
  mock.read.mockImplementationOnce(() => new Promise((resolve) => { finish = resolve }))
  const selected = vi.fn()
  const stop = watchSelectedDiagram(selected, vi.fn())
  stop()
  finish(null)
  stop.refresh()
  mock.change()
  await new Promise((resolve) => setTimeout(resolve, 0))
  expect(selected).not.toHaveBeenCalled()
  expect(mock.read).toHaveBeenCalledOnce()
  expect(mock.remove).toHaveBeenCalledOnce()
})
