import {
  DIAGRAM_THEMES,
  type DiagramTheme,
} from '../metadata/payload'
import { DEFAULT_DIAGRAM_SETTINGS, isDiagramSettings, type DiagramSettings } from '../metadata/diagramSettings'

const THEME_KEY = 'mermaid-office:preferred-theme'
const SETTINGS_KEY = 'mermaid-office:preferred-settings'

function readPreference<T extends string>(
  key: string,
  allowed: readonly T[],
  fallback: T,
): T {
  if (typeof window === 'undefined') {
    return fallback
  }
  try {
    const value = window.localStorage.getItem(key) as T | null
    return value && allowed.includes(value) ? value : fallback
  } catch (error) {
    console.warn(`Unable to read the ${key} preference.`, error)
    return fallback
  }
}

function writePreference(key: string, value: string): void {
  try {
    window.localStorage.setItem(key, value)
  } catch (error) {
    console.warn(`Unable to save the ${key} preference.`, error)
  }
}

export function getPreferredTheme(): DiagramTheme {
  return readPreference(THEME_KEY, DIAGRAM_THEMES, 'redux-color')
}

export function setPreferredTheme(theme: DiagramTheme): void {
  writePreference(THEME_KEY, theme)
}

export function getPreferredSettings(): DiagramSettings {
  if (typeof window === 'undefined') return DEFAULT_DIAGRAM_SETTINGS
  try {
    const value = window.localStorage.getItem(SETTINGS_KEY)
    if (value === null) return DEFAULT_DIAGRAM_SETTINGS
    const settings: unknown = JSON.parse(value)
    if (!isDiagramSettings(settings)) throw new Error('Unsupported diagram settings preference.')
    return settings
  } catch (error) {
    console.warn('Unable to read the diagram settings preference.', error)
    return DEFAULT_DIAGRAM_SETTINGS
  }
}

export function setPreferredSettings(settings: DiagramSettings): void {
  writePreference(SETTINGS_KEY, JSON.stringify(settings))
}
