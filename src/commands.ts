import { openEditorDialog } from './dialog/openEditorDialog'
import { DEFAULT_DIAGRAM } from './defaultDiagram'
import { renderMermaid } from './mermaid/render'
import { insertDiagram, updateDiagram } from './word/insertDiagram'
import { getSelectedDiagram } from './word/selection'
import {
  getPreferredSize,
  getPreferredTheme,
  setPreferredSize,
  setPreferredTheme,
} from './preferences/diagramPreferences'

async function insertDefaultDiagram(event: Office.AddinCommands.Event) {
  try {
    const theme = getPreferredTheme()
    const size = getPreferredSize()
    const svg = await renderMermaid(DEFAULT_DIAGRAM, theme)
    await insertDiagram(svg, DEFAULT_DIAGRAM, theme, size)
  } catch (error) {
    console.error('Unable to insert the default Mermaid diagram.', error)
  } finally {
    event.completed()
  }
}

async function openEditor(event: Office.AddinCommands.Event) {
  try {
    const selectedDiagram = await getSelectedDiagram()
    const result = await openEditorDialog(
      selectedDiagram?.source ?? DEFAULT_DIAGRAM,
      selectedDiagram?.theme ?? getPreferredTheme(),
      selectedDiagram?.size ?? getPreferredSize(),
    )
    if (result === null) {
      return
    }

    const svg = await renderMermaid(result.source, result.theme)
    setPreferredTheme(result.theme)
    setPreferredSize(result.size)
    if (selectedDiagram) {
      await updateDiagram(
        svg,
        selectedDiagram,
        result.source,
        result.theme,
        result.size,
        result.size !== selectedDiagram.size,
      )
    } else {
      await insertDiagram(svg, result.source, result.theme, result.size)
    }
  } catch (error) {
    console.error('Unable to edit the Mermaid diagram.', error)
  } finally {
    event.completed()
  }
}

Office.onReady(() => {
  Office.actions.associate('insertDefaultDiagram', insertDefaultDiagram)
  Office.actions.associate('openEditor', openEditor)
})
