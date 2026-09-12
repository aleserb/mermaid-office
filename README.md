# Mermaid Office

Mermaid Office is a client-side Microsoft Word add-in for creating and editing
Mermaid diagrams. Mermaid renders locally in the Office task pane; no
application server receives diagram source or document content.

See [PUBLISHING.md](PUBLISHING.md) for Microsoft Marketplace listing information,
public policy URLs, certification instructions, and the submission checklist.

## License, privacy, and support

The project's original code and documentation are licensed under the
[MIT License](LICENSE). Third-party components retain their own licenses.
Production builds include generated
[third-party license notices](https://aleserb.github.io/mermaid-office/third-party-licenses.txt).

- [Privacy policy](https://aleserb.github.io/mermaid-office/privacy.html)
- [End-user license agreement (EULA)](https://aleserb.github.io/mermaid-office/eula.html)
- [Support and troubleshooting](https://aleserb.github.io/mermaid-office/support.html)

These standalone pages are served from `public/` and do not require Word,
Office.js, or an account to read. Diagram source is embedded in original diagram
images and document metadata; sharing those files may also share the source.

## Technology

- React, TypeScript, and Vite
- Office.js with client-side PNG insertion
- Mermaid.js and DOMPurify
- CodeMirror 6
- Fluent UI React v9
- Vitest and Testing Library
- GitHub Pages

The add-in creates a tightly cropped PNG locally before inserting it into Word.
Using one image format across desktop and web avoids host-specific SVG sizing
differences and keeps insertion and update behavior consistent.
The final PNG is rasterized directly from SVG. Auto quality (the default) targets
2 pixels per displayed CSS pixel for normal viewing on a high-density screen,
retaining native SVG detail for dense diagrams when the budget permits.
New diagrams use the fitted Word size; edits use the picture's actual width,
including manual resizing, rather than always targeting extreme zoom.
Resolution is recalculated on insertion or a valid edit, not on Word zoom changes
or resizing alone. Standard quality uses the same display target without retaining
extra native detail for dense diagrams. Auto and Standard exports are bounded to
4096 pixels per side and about 4.2 megapixels. High quality targets 3x native detail
or 8x the displayed CSS width, bounded to 8192 pixels per side and about 33.6 megapixels.
High quality uses more memory and produces larger PNGs;
canvas backing stores are released immediately after encoding.
Cropping, displayed Word size, and embedded source metadata are preserved.
Small diagrams use Word's native PNG insertion and replacement API, avoiding the
OOXML import that triggers Word's blocking "Waiting..." dialog even for tiny images.
This path is limited to PNGs at most 1536 pixels per side and 2 Mi-pixels, in frames
at most 468 points wide and 432 points tall. Source metadata, alternative text,
and the displayed width are preserved on updates.
Larger or taller pictures use an inline OOXML drawing with explicit frame and image
dimensions: native replacement can auto-fit these pictures or paint outside the frame.
Word on the web can still show its own blocking "Waiting..." dialog when importing
these larger drawings, including during live updates. This is a Word-owned dialog,
not an add-in editor dialog; Office.js provides no supported switch to hide it.
PNG remains a raster format, so extreme zoom can still reveal pixels. Existing
images adopt the automatic resolution and insertion path on their next valid edit.

## Development

```bash
npm install
npm run dev
```

The regular browser view opens the code-only pane for UI development. Diagram
insertion must be tested by sideloading `manifest.xml` in Microsoft Word.
After manifest changes, remove and upload the manifest again because Word caches
the sideloaded manifest separately from the hosted web application.

When the task pane is open, selecting an inserted Mermaid diagram loads its
stored source and theme into the editor automatically. Valid edits replace the
selected picture in place while preserving its displayed width and alternative
text. Large diagrams use explicit updates instead, as described below.
Move the cursor away from the diagram to leave edit mode.
Diagram wrappers use a blank placeholder, so deleting a picture does not leave
Word's "Click or tap here to enter text" prompt, even with the add-in closed.
Older diagrams receive this setting when selected or updated. With the pane open,
placing the cursor in a leftover diagram wrapper removes the wrapper while
preserving any surrounding text. Saved metadata is retained for Word Undo.

When Word preserves the original PNG or SVG bytes during copy/paste, selecting
the pasted diagram restores its editable Mermaid source from embedded metadata.
Copied content controls with duplicate IDs are assigned a new identity before
editing so changes stay attached to the selected diagram.

The Word picture itself is the live preview. Parse failures appear below the
code editor and in its gutter, keep the last valid picture visible in Word,
and prevent insertion or live updates until the source is valid.

The source editor includes Mermaid-aware syntax highlighting and autocomplete,
common diagram snippets, automatic bracket and quote closing, standard editor
keyboard shortcuts, and a link to the Mermaid syntax reference.
The fixed-height header places Insert on the left and the settings gear and
syntax-reference document-info icon on the right. Insert is hidden while editing
an existing diagram without shifting the editor. Large diagrams show Update in
the same space; it is disabled until valid, unsaved changes are ready.
The active line and its line number are highlighted. The editor's built-in
search and replace panel and its shortcuts are disabled.
In sequence diagrams, autocomplete also suggests participant and actor IDs from
declarations and existing messages, with declared aliases shown as descriptions.
Suggestions are available for message senders and recipients, note references,
and activation commands; use Ctrl+Space to request suggestions explicitly.

The editor UI always uses a light appearance, regardless of
Word's or the system's theme. The selected Mermaid theme independently controls
the diagram's colors and exported PNG.
The code editor always uses the platform's System UI font at 12 px.
This changes only the editor display, not the diagram output.

The gear button opens an Office-hosted settings window outside the task pane.
In Word on the web it floats over the Word workspace; desktop Word uses a separate,
centered window. The code editor stays in the pane. Browser-only development, or
hosts without the Office Dialog API, retain the in-pane settings form.
Theme is in this window rather than above the code editor.
Use Apply to change all selected settings together, or Cancel/Escape to discard
the dialog changes. The window is non-modal, so the Word document remains
interactive. The pane is temporarily locked to the original diagram: Apply updates
that diagram before following a new Word selection. Cancel follows the new
selection without applying settings. Deleted or externally changed diagrams
are rejected by the existing update guards.
Office-hosted settings require DialogApi 1.2. Loading, popup permission, and messaging
errors are shown in the pane instead of silently reverting to a narrow dialog.

Settings offers Mermaid's Default, Neutral, Dark, Forest, Neo, Redux Color,
and monochrome Redux render themes, including their dark variants. The selected
theme is stored with the diagram so it is restored when editing later. The most
recently selected theme is also remembered for new diagrams. New users start
with Redux Color when no theme preference has been saved.
Additional settings include diagram font and font size, visual style, flowchart
spacing/connectors/layout, sequence spacing/numbering/wrapping/bottom participants,
and Auto/Standard/High PNG quality. Type-specific controls appear only for the
relevant diagram. There is no custom-color editor.
ELK flowchart layout loads an additional bundled client-side module only when needed;
no diagram data is sent to a layout service.
These settings are stored with the diagram in Word and embedded PNG metadata.
The last applied settings are also remembered for new diagrams; older diagrams
without settings retain their original defaults. Diagram font settings do not
change the code editor's System UI font.
High PNG quality can switch even a small diagram to the OOXML insertion path and
cause Word's "Waiting..." dialog; Auto is recommended for routine live editing.
Default-valued options retain existing source configuration. Resetting an option
to its default does not suppress explicit frontmatter/init settings or directives
such as `autonumber`. Visual-look support varies by diagram type; sequence diagrams
do not support Hand-drawn.

The ribbon has one **Insert > Mermaid** button, which opens the code-only
right-hand task pane. There is no separate editor dialog or preview pane.
Remove and upload the updated `manifest.xml` again to replace the old ribbon
buttons; reopening a cached pane alone does not refresh Word's ribbon.
Place the cursor on a blank line and press **Insert** once. For small diagrams,
valid edits then update the PNG directly in Word after a short typing pause.
For large diagrams, edits remain in the pane until **Update** is pressed.
Validation still runs while typing; invalid source leaves the last valid image
in place and disables Update. Settings **Apply** is also an explicit update,
including any pending source edits.

Manual updates are enabled at 50 nonblank source lines, 4,000 source characters,
or rendered SVG dimensions exceeding 1,536 pixels on either side or 2 megapixels
(2,097,152 pixels). These are conservative size heuristics, not an exact prediction
of Word's image-import behavior. Once enabled, manual mode stays active for the
current editing session, even if the diagram is shortened; selecting another
diagram resets the mode and evaluates its size again. Edits made during an Update
remain unsaved until the next click. Switching diagrams with unsaved changes still
requires confirmation. This avoids repeated large-picture imports while typing;
an explicit Update can still show Word's own loading dialog.

The pane automatically follows Word's selection. Select a Mermaid picture to
load its source, or move to ordinary text or a blank line to start with the
default source and **Insert** button. Repeated selection notifications
do not reset your current draft. Switching with pending edits requires confirmation.
Live updates modify the real document and can add Word undo/AutoSave changes;
discarding pending edits does not undo changes already written to Word.

## Checks

```bash
npm run lint
npm test
npm run build
```

## Deployment

Run `npm run deploy` to build the application and publish `dist` to the
`gh-pages` branch. Configure the repository's Pages source as the root of that
branch before the first deployment.

GitHub Pages for a private repository requires a GitHub plan that supports
private Pages sites. Otherwise, make the repository public or upgrade the
account before enabling Pages.
