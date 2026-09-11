import {
  DIAGRAM_THEMES,
  type DiagramTheme,
} from '../metadata/payload'

const THEME_KEY = 'mermaid-office:preferred-theme'

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
  return readPreference(THEME_KEY, DIAGRAM_THEMES, 'default')
}

export function setPreferredTheme(theme: DiagramTheme): void {
  writePreference(THEME_KEY, theme)
}
