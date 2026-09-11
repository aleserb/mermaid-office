import { describe, expect, it } from 'vitest'
import { parseDialogMessage, parseParentMessage } from './messages'

describe('editor dialog messages', () => {
  it('parses initialization and save messages', () => {
    expect(
      parseParentMessage(
        '{"type":"initialize","source":"flowchart LR","theme":"redux-color","size":"large"}',
      ),
    ).toEqual({
      type: 'initialize',
      source: 'flowchart LR',
      theme: 'redux-color',
      size: 'large',
    })
    expect(
      parseDialogMessage(
        '{"type":"save","source":"sequenceDiagram","theme":"forest","size":"small"}',
      ),
    ).toEqual({
      type: 'save',
      source: 'sequenceDiagram',
      theme: 'forest',
      size: 'small',
    })
  })

  it('rejects malformed messages', () => {
    expect(() => parseParentMessage('{"type":"initialize"}')).toThrow('invalid')
    expect(() => parseDialogMessage('{"type":"unknown"}')).toThrow('unsupported')
  })
})
