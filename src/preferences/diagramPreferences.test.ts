import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  getPreferredSize,
  getPreferredTheme,
  setPreferredSize,
  setPreferredTheme,
} from './diagramPreferences'

describe('diagram preferences', () => {
  afterEach(() => {
    window.localStorage.clear()
    vi.restoreAllMocks()
  })

  it('persists valid theme and size values', () => {
    setPreferredTheme('redux-color')
    setPreferredSize('large')

    expect(getPreferredTheme()).toBe('redux-color')
    expect(getPreferredSize()).toBe('large')
  })

  it('uses safe defaults when storage is unavailable', () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new DOMException('Blocked', 'SecurityError')
    })
    vi.spyOn(console, 'warn').mockImplementation(() => undefined)

    expect(getPreferredTheme()).toBe('default')
    expect(getPreferredSize()).toBe('medium')
  })
})
