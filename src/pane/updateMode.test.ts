import { expect, it } from 'vitest'
import { isLargeDiagram } from './updateMode'

it('uses nonblank source lines and character count as early size signals', () => {
  expect(isLargeDiagram('A\n'.repeat(49))).toBe(false)
  expect(isLargeDiagram('A\r\n'.repeat(50))).toBe(true)
  expect(isLargeDiagram('\n \n'.repeat(100))).toBe(false)
  expect(isLargeDiagram('A'.repeat(3999))).toBe(false)
  expect(isLargeDiagram('A'.repeat(4000))).toBe(true)
})

it.each([
  [300, 100, false],
  [1536, 100, false],
  [1537, 100, true],
  [100, 1537, true],
  [1500, 1500, true],
])('classifies a %s by %s SVG as large=%s', (width, height, large) => {
  expect(isLargeDiagram('flowchart LR\nA-->B', `<svg viewBox="0 0 ${width} ${height}"/>`)).toBe(large)
  expect(isLargeDiagram('flowchart LR\nA-->B', `<svg width="${width}" height="${height}"/>`)).toBe(large)
})
