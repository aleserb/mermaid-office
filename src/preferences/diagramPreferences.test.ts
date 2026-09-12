import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  getPreferredTheme,
  setPreferredTheme,
} from './diagramPreferences'
import { DIAGRAM_THEMES } from '../metadata/payload'

describe('diagram preferences', () => {
  afterEach(() => {
    window.localStorage.clear()
    vi.restoreAllMocks()
  })

  it('defaults to Redux Color for a new user', () => {
    expect(getPreferredTheme()).toBe('redux-color')
  })

  it.each(DIAGRAM_THEMES)('preserves the saved %s theme', (theme) => {
    setPreferredTheme(theme)

    expect(getPreferredTheme()).toBe(theme)
  })

  it('uses Redux Color when the saved preference is invalid', () => {
    window.localStorage.setItem('mermaid-office:preferred-theme', 'unknown')

    expect(getPreferredTheme()).toBe('redux-color')
  })

  it('uses safe defaults when storage is unavailable', () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new DOMException('Blocked', 'SecurityError')
    })
    vi.spyOn(console, 'warn').mockImplementation(() => undefined)

    expect(getPreferredTheme()).toBe('redux-color')
  })
})
