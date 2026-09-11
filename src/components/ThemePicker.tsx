import type { DiagramTheme } from '../metadata/payload'
import './ThemePicker.css'

const themes: Array<{ value: DiagramTheme; label: string }> = [
  { value: 'default', label: 'Default' },
  { value: 'neutral', label: 'Neutral' },
  { value: 'dark', label: 'Dark' },
  { value: 'forest', label: 'Forest' },
  { value: 'neo', label: 'Neo' },
  { value: 'neo-dark', label: 'Neo Dark' },
  { value: 'redux-color', label: 'Redux Color' },
  { value: 'redux-dark-color', label: 'Redux Dark Color' },
  { value: 'redux', label: 'Redux Monochrome' },
  { value: 'redux-dark', label: 'Redux Dark Monochrome' },
]

interface ThemePickerProps {
  value: DiagramTheme
  onChange: (theme: DiagramTheme) => void
}

export function ThemePicker({ value, onChange }: ThemePickerProps) {
  return (
    <label className="theme-picker">
      <span>Theme</span>
      <select
        aria-label="Diagram theme"
        value={value}
        onChange={(event) => onChange(event.target.value as DiagramTheme)}
      >
        {themes.map((theme) => (
          <option key={theme.value} value={theme.value}>
            {theme.label}
          </option>
        ))}
      </select>
    </label>
  )
}
