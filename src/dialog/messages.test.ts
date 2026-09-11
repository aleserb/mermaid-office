import { describe, expect, it } from 'vitest'
import { parseDialogMessage, parseParentMessage } from './messages'

describe('editor dialog messages', () => {
  it('parses initialization and save messages', () => {
    expect(parseParentMessage('{"type":"initialize","source":"flowchart LR"}')).toEqual({
      type: 'initialize',
      source: 'flowchart LR',
    })
    expect(parseDialogMessage('{"type":"save","source":"sequenceDiagram"}')).toEqual({
      type: 'save',
      source: 'sequenceDiagram',
    })
  })

  it('rejects malformed messages', () => {
    expect(() => parseParentMessage('{"type":"initialize"}')).toThrow('invalid')
    expect(() => parseDialogMessage('{"type":"unknown"}')).toThrow('unsupported')
  })
})
