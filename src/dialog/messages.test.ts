import { describe, expect, it } from 'vitest'
import { parseDialogMessage, parseParentMessage } from './messages'

describe('editor dialog messages', () => {
  it('parses initialization and save messages', () => {
    expect(
      parseParentMessage(
        '{"type":"initialize","source":"flowchart LR","theme":"redux"}',
      ),
    ).toEqual({
      type: 'initialize',
      source: 'flowchart LR',
      theme: 'redux',
    })
    expect(
      parseDialogMessage(
        '{"type":"save","source":"sequenceDiagram","theme":"forest"}',
      ),
    ).toEqual({
      type: 'save',
      source: 'sequenceDiagram',
      theme: 'forest',
    })
  })

  it('rejects malformed messages', () => {
    expect(() => parseParentMessage('{"type":"initialize"}')).toThrow('invalid')
    expect(() => parseDialogMessage('{"type":"unknown"}')).toThrow('unsupported')
  })
})
