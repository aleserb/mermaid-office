# Mermaid Office

Mermaid Office is a client-side Microsoft Word add-in for creating and editing
Mermaid diagrams. Mermaid renders locally in the Office task pane; no
application server receives diagram source or document content.

See [design.md](design.md) for the product goals, user experience, architecture,
compatibility strategy, and delivery plan.

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
The final PNG is rasterized directly from SVG, targeting 2x the diagram's native
dimensions or 4x its intended display width, whichever provides more detail.
Exports are bounded to 8192 pixels per side and about 16.8 megapixels to limit memory use.
Cropping, displayed Word size, and embedded source metadata are preserved.
Pictures are inserted and updated as an inline Word drawing with explicit frame
and image dimensions, keeping higher-resolution pixels inside the picture frame.
PNG remains a raster format, so extreme zoom can still reveal pixels. Existing
images gain the higher resolution when a valid edit regenerates them.

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
text. Move the cursor away from the diagram to leave edit mode.
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

The editor offers Mermaid's Default, Neutral, Dark, Forest, Neo, Redux Color,
and monochrome Redux render themes, including their dark variants. The selected
theme is stored with the diagram so it is restored when editing later. The most
recently selected theme is also remembered for new diagrams. New users start
with Redux Color when no theme preference has been saved.

The ribbon has one **Insert > Mermaid** button, which opens the code-only
right-hand task pane. There is no separate editor dialog or preview pane.
Remove and upload the updated `manifest.xml` again to replace the old ribbon
buttons; reopening a cached pane alone does not refresh Word's ribbon.
Place the cursor on a blank line and press **Insert** once. Valid source
and theme edits then update that diagram's PNG directly in Word after a short
typing pause; invalid source leaves the last valid image in place.

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
