import { openEditorDialog } from './dialog/openEditorDialog'
import { DEFAULT_DIAGRAM } from './defaultDiagram'
import { renderMermaid } from './mermaid/render'
import { insertDiagram, updateDiagram } from './word/insertDiagram'
import { getSelectedDiagram } from './word/selection'

async function insertDefaultDiagram(event: Office.AddinCommands.Event) {
  try {
    const svg = await renderMermaid(DEFAULT_DIAGRAM)
    await insertDiagram(svg, DEFAULT_DIAGRAM)
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
      selectedDiagram?.theme ?? 'default',
    )
    if (result === null) {
      return
    }

    const svg = await renderMermaid(result.source, result.theme)
    if (selectedDiagram) {
      await updateDiagram(svg, selectedDiagram, result.source, result.theme)
    } else {
      await insertDiagram(svg, result.source, result.theme)
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
