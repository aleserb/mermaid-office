import { getContentControlTag, type DiagramPayload } from '../metadata/payload'

export function suppressDiagramPlaceholder(contentControl: Word.ContentControl): void {
  // An empty string restores Word's default prompt; a space keeps it invisible
  // even when the picture is deleted while the add-in is closed.
  contentControl.placeholderText = ' '
}

export function configureDiagramContentControl(
  contentControl: Word.ContentControl,
  payload: Pick<DiagramPayload, 'id'>,
): void {
  contentControl.tag = getContentControlTag(payload.id)
  contentControl.title = 'Mermaid diagram'
  contentControl.appearance = Word.ContentControlAppearance.hidden
  contentControl.cannotDelete = false
  contentControl.cannotEdit = false
  suppressDiagramPlaceholder(contentControl)
}
