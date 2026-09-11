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
    const source = await openEditorDialog(selectedDiagram?.source ?? DEFAULT_DIAGRAM)
    if (source === null) {
      return
    }

    const svg = await renderMermaid(source)
    if (selectedDiagram) {
      await updateDiagram(svg, selectedDiagram, source)
    } else {
      await insertDiagram(svg, source)
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
