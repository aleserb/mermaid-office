export interface EditorDisplay {
  fontFamily: string
  fontSize: number
}

export const EDITOR_FONTS = [
  { label: 'Default mono', value: '"SFMono-Regular", Consolas, "Liberation Mono", monospace' },
  { label: 'System UI', value: 'system-ui, sans-serif' },
  { label: 'Segoe UI', value: '"Segoe UI", system-ui, sans-serif' },
  { label: 'Consolas', value: 'Consolas, monospace' },
  { label: 'Cascadia Code', value: '"Cascadia Code", monospace' },
  { label: 'Menlo', value: 'Menlo, monospace' },
  { label: 'Courier New', value: '"Courier New", monospace' },
  { label: 'Arial', value: 'Arial, sans-serif' },
]

export const EDITOR_FONT_SIZES = [10, 12, 13, 14, 16, 18, 20, 24]

export const DEFAULT_EDITOR_DISPLAY: EditorDisplay = {
  fontFamily: EDITOR_FONTS[0].value,
  fontSize: 13,
}
