import { openEditorDialog } from './dialog/openEditorDialog'
import { DEFAULT_DIAGRAM } from './defaultDiagram'
import { insertDiagram, updateDiagram } from './word/insertDiagram'
import { getSelectedDiagram } from './word/selection'
import {
  getPreferredTheme,
  setPreferredTheme,
} from './preferences/diagramPreferences'

async function openEditor(event: Office.AddinCommands.Event) {
  try {
    const selectedDiagram = await getSelectedDiagram()
    const result = await openEditorDialog(
      selectedDiagram?.source ?? DEFAULT_DIAGRAM,
      selectedDiagram?.theme ?? getPreferredTheme(),
      selectedDiagram?.size ?? 'medium',
      selectedDiagram ? 'update' : 'insert',
    )
    if (result === null) {
      return
    }

    setPreferredTheme(result.theme)
    if (selectedDiagram) {
      await updateDiagram(
        result.svg,
        selectedDiagram,
        result.source,
        result.theme,
        result.size,
        result.size !== selectedDiagram.size,
        result.raster,
      )
    } else {
      await insertDiagram(
        result.svg,
        result.source,
        result.theme,
        result.size,
        result.raster,
      )
    }
  } catch (error) {
    console.error('Unable to edit the Mermaid diagram.', error)
  } finally {
    event.completed()
  }
}

Office.onReady(() => {
  Office.actions.associate('openEditor', openEditor)
  // Keep cached copies of the previous manifest functional until Word refreshes them.
  Office.actions.associate('insertDefaultDiagram', openEditor)
})
