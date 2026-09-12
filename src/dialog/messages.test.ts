import { describe, expect, it } from 'vitest'
import { parseDialogMessage, parseParentMessage } from './messages'

describe('editor dialog messages', () => {
  it('parses initialization and save messages', () => {
    expect(
      parseParentMessage(
        '{"type":"initialize","mode":"update","source":"flowchart LR","theme":"redux-color","size":"large"}',
      ),
    ).toEqual({
      type: 'initialize',
      mode: 'update',
      source: 'flowchart LR',
      theme: 'redux-color',
      size: 'large',
    })
    expect(
      parseDialogMessage(
        '{"type":"save","source":"sequenceDiagram","theme":"forest","size":"small","svg":"<svg></svg>","raster":{"base64":"png","width":320,"height":180}}',
      ),
    ).toEqual({
      type: 'save',
      source: 'sequenceDiagram',
      theme: 'forest',
      size: 'small',
      svg: '<svg></svg>',
      raster: { base64: 'png', width: 320, height: 180 },
    })
  })

  it('rejects malformed messages', () => {
    expect(() => parseParentMessage('{"type":"initialize"}')).toThrow('invalid')
    expect(() => parseDialogMessage('{"type":"unknown"}')).toThrow('unsupported')
  })

  it('accepts initialization from cached runtimes with an unused UI theme flag', () => {
    const initialization = {
      type: 'initialize',
      mode: 'insert',
      source: 'sequenceDiagram',
      theme: 'forest',
      size: 'medium',
    }
    expect(parseParentMessage(JSON.stringify({ ...initialization, darkMode: true })))
      .toMatchObject(initialization)
  })
})
