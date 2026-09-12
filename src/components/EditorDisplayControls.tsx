import { EDITOR_FONTS, EDITOR_FONT_SIZES, type EditorDisplay } from './editorDisplay'
import './EditorDisplayControls.css'

interface EditorDisplayControlsProps {
  value: EditorDisplay
  onChange: (value: EditorDisplay) => void
}

export function EditorDisplayControls({ value, onChange }: EditorDisplayControlsProps) {
  return (
    <div className="editor-display-controls" role="group" aria-label="Code editor display">
      <label>
        Font
        <select
          aria-label="Code editor font"
          title="Uses installed fonts; unavailable fonts use a system fallback."
          value={value.fontFamily}
          onChange={(event) => onChange({ ...value, fontFamily: event.target.value })}
        >
          {EDITOR_FONTS.map((font) => (
            <option key={font.value} value={font.value}>{font.label}</option>
          ))}
        </select>
      </label>
      <label>
        Size
        <select
          aria-label="Code editor font size"
          value={value.fontSize}
          onChange={(event) => onChange({ ...value, fontSize: Number(event.target.value) })}
        >
          {EDITOR_FONT_SIZES.map((size) => (
            <option key={size} value={size}>{size} px</option>
          ))}
        </select>
      </label>
    </div>
  )
}
