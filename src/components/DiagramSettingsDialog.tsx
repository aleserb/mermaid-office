import { useId, useState, type ReactElement } from 'react'
import {
  Button,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  Field,
  Select,
} from '@fluentui/react-components'
import type { DiagramTheme } from '../metadata/payload'
import type { DiagramKind, DiagramSettings } from '../metadata/diagramSettings'
import { ThemePicker } from './ThemePicker'
import './DiagramSettingsDialog.css'

interface DiagramSettingsDialogProps {
  open?: boolean
  trigger?: ReactElement
  onOpen?: () => void
  theme: DiagramTheme
  settings: DiagramSettings
  diagramKind: DiagramKind
  disabled?: boolean
  onCancel: () => void
  onApply: (theme: DiagramTheme, settings: DiagramSettings) => void
}

export function DiagramSettingsDialog({ open = true, trigger, onOpen, ...props }: DiagramSettingsDialogProps) {
  const [session, setSession] = useState({ open, revision: 0 })
  if (session.open !== open) {
    setSession({ open, revision: session.revision + (open ? 1 : 0) })
  }
  const surface = (
    <DialogSurface key="surface" className="diagram-settings-surface">
      <DiagramSettingsForm key={session.revision} {...props} disabled={props.disabled || !open} />
    </DialogSurface>
  )
  const children: ReactElement | [ReactElement, ReactElement] = trigger
    ? [<DialogTrigger key="trigger" action="open" disableButtonEnhancement>{trigger}</DialogTrigger>, surface]
    : surface
  // Keep Dialog mounted so Fluent restores background accessibility on close.
  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (data.open) onOpen?.(); else props.onCancel() }}>
      {children}
    </Dialog>
  )
}

function DiagramSettingsForm({
  theme,
  settings,
  diagramKind,
  disabled = false,
  onCancel,
  onApply,
}: Omit<DiagramSettingsDialogProps, 'open' | 'trigger' | 'onOpen'>) {
  const [draftTheme, setDraftTheme] = useState(theme)
  const [draftSettings, setDraftSettings] = useState(() => ({ ...settings }))
  const qualityDescriptionId = useId()
  const lookDescriptionId = useId()
  const change = <K extends keyof DiagramSettings>(key: K, value: DiagramSettings[K]) => {
    setDraftSettings(current => ({ ...current, [key]: value }))
  }

  return (
    <DialogBody className="diagram-settings-body">
      <DialogTitle>Diagram settings</DialogTitle>
      <DialogContent className="diagram-settings-content">
        <p className="diagram-settings-note">
          Settings affect the diagram, not the code editor, and are saved when applied.
        </p>
        <ThemePicker value={draftTheme} onChange={setDraftTheme} />
        <Field label="Font family">
          <Select value={draftSettings.fontFamily} onChange={(_, data) =>
            change('fontFamily', data.value as DiagramSettings['fontFamily'])}>
            <option value="theme">Theme default</option>
            <option value="system-ui">System UI</option>
            <option value="Arial">Arial</option>
            <option value="Times New Roman">Times New Roman</option>
            <option value="monospace">Monospace</option>
          </Select>
        </Field>
        <Field label="Font size">
          <Select value={String(draftSettings.fontSize)} onChange={(_, data) =>
            change('fontSize', data.value === 'theme' ? 'theme' : Number(data.value) as DiagramSettings['fontSize'])}>
            <option value="theme">Theme default</option>
            {[12, 14, 16, 18, 20, 24].map(size => <option key={size} value={size}>{size}</option>)}
          </Select>
        </Field>
        <Field label="Visual look">
          <Select value={draftSettings.look} aria-describedby={lookDescriptionId} onChange={(_, data) =>
            change('look', data.value as DiagramSettings['look'])}>
            <option value="auto">Theme default</option>
            <option value="classic">Classic</option>
            <option value="neo">Neo</option>
            <option value="handDrawn" disabled={diagramKind === 'sequence'}>Hand-drawn</option>
          </Select>
        </Field>
        <p className="diagram-settings-note" id={lookDescriptionId}>
          Visual look support varies by diagram type.
          {diagramKind === 'sequence' &&
            ' Hand-drawn is not supported for sequence diagrams. A saved Hand-drawn choice is retained until you change it.'}
        </p>
        {(diagramKind === 'flowchart' || diagramKind === 'sequence') && (
          <Field label="Spacing">
            <Select value={draftSettings.spacing} onChange={(_, data) =>
              change('spacing', data.value as DiagramSettings['spacing'])}>
              <option value="default">Default</option>
              <option value="compact">Compact</option>
              <option value="spacious">Spacious</option>
            </Select>
          </Field>
        )}
        {diagramKind === 'flowchart' && <>
          <Field label="Connector style">
            <Select value={draftSettings.curve} onChange={(_, data) =>
              change('curve', data.value as DiagramSettings['curve'])}>
              <option value="default">Default</option>
              <option value="linear">Straight</option>
              <option value="basis">Curved</option>
              <option value="step">Stepped</option>
            </Select>
          </Field>
          <Field label="Layout">
            <Select value={draftSettings.layout} onChange={(_, data) =>
              change('layout', data.value as DiagramSettings['layout'])}>
              <option value="default">Default</option>
              <option value="dagre">Dagre</option>
              <option value="elk">ELK</option>
            </Select>
          </Field>
        </>}
        {diagramKind === 'sequence' && (
          <div className="diagram-settings-toggles">
            <Checkbox label="Number messages" checked={draftSettings.sequenceNumbers}
              onChange={(_, data) => change('sequenceNumbers', data.checked === true)} />
            <Checkbox label="Wrap text" checked={draftSettings.sequenceWrap}
              onChange={(_, data) => change('sequenceWrap', data.checked === true)} />
            <Checkbox label="Repeat participants below" checked={draftSettings.sequenceMirrorActors}
              onChange={(_, data) => change('sequenceMirrorActors', data.checked === true)} />
            <p className="diagram-settings-note">
              Source configuration and directives such as autonumber can control default-valued options.
            </p>
          </div>
        )}
        <Field label="Image quality">
          <Select value={draftSettings.imageQuality} aria-describedby={qualityDescriptionId} onChange={(_, data) =>
            change('imageQuality', data.value as DiagramSettings['imageQuality'])}>
            <option value="auto">Auto (adaptive)</option>
            <option value="standard">Standard (smaller PNGs)</option>
            <option value="high">High (sharper zoom)</option>
          </Select>
        </Field>
        <p className="diagram-settings-note" id={qualityDescriptionId}>
          Image quality controls the output PNG, not the editor font.
          {draftSettings.imageQuality === 'high' &&
            " High quality creates larger, slower PNGs and may trigger Word's loading dialog."}
        </p>
      </DialogContent>
      <DialogActions className="diagram-settings-actions">
        <Button onClick={onCancel}>Cancel</Button>
        <Button appearance="primary" disabled={disabled}
          onClick={() => { if (!disabled) onApply(draftTheme, { ...draftSettings }) }}>
          Apply
        </Button>
      </DialogActions>
    </DialogBody>
  )
}
