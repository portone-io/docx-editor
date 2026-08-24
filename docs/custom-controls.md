# Custom controls

`DocxEditor` can keep its editing surface while a host application replaces the built-in toolbar and context menus.

## Hide the built-in controls

Disable either surface through `mode`:

```tsx
<DocxEditor
  document={file}
  mode={{ kind: "edit", toolbar: false, contextMenus: false }}
/>
```

## Document zoom

The editor fits the page to its available width by default. Its toolbar offers Fit, 50%, 75%, 100%, 125%, and 150%. Fixed zoom remains selected on a narrow screen and makes the document horizontally scrollable instead of forcing it back to fit.

Use `defaultZoom` when the editor should own the selection, or `zoom` with `onZoomChange` when the host should own it:

```tsx
import { DocxEditor, type DocxEditorZoom } from "@portone-io/docx-editor";
import { useState } from "react";

function Preview({ file }: { file: File }) {
  const [zoom, setZoom] = useState<DocxEditorZoom>("fit-width");

  return (
    <DocxEditor
      document={file}
      zoom={zoom}
      onZoomChange={setZoom}
    />
  );
}
```

Numeric values from `0.25` through `2` are accepted. The root entry exports `DEFAULT_ZOOM_LEVELS` for a custom toolbar. A read-only editor has no built-in toolbar, but a host can still set `zoom`. Zoom changes only the browser presentation; it does not change pagination measurements, the document model, or exported OOXML. Persisting a reader's choice belongs to the host application.

The comment panel reads the same factor from the `--docx-editor-zoom` custom property, set on the editor workspace, so comment text scales with the paper instead of standing apart from it. Its font sizes stop shrinking at a readable floor, and the panel keeps its width at every zoom level.

## Commands and queries

Text and paragraph commands come from `@portone-io/docx-editor/commands`. Operations on an existing table come from `@portone-io/docx-editor/table`, while `insertTable` comes from `@portone-io/docx-editor/commands`.

Commands follow the ProseMirror command contract. Calling one with an `EditorState` reports whether it applies, while passing a dispatch function performs the edit.

```tsx
import type {
  Command,
  EditorState,
  EditorView,
} from "@portone-io/docx-editor";
import {
  isBoldActive,
  toggleBold,
  undo,
} from "@portone-io/docx-editor/commands";

function Toolbar({ view, state }: { view: EditorView; state: EditorState }) {
  const run = (command: Command) => {
    command(state, (transaction) => view.dispatch(transaction), view);
    view.focus();
  };

  return (
    <div>
      <button
        type="button"
        aria-pressed={isBoldActive(state)}
        onClick={() => run(toggleBold)}
      >
        Bold
      </button>
      <button type="button" disabled={!undo(state)} onClick={() => run(undo)}>
        Undo
      </button>
    </div>
  );
}
```

`onReady` provides the editor view, and `onChange` fires whenever the editor state changes. Use both to keep custom controls in sync with the current selection.

The query prefixes describe what they return:

- `is*Active` reports a boolean toggle.
- `active*` reports the current value or a mixed selection.
- `can*` reports whether the related action applies.
- `isIn*` reports containment.

`DEFAULT_X` names the built-in default for the prop or field named `x`. Selection-level readers use `selection*`, while document-level readers use `document*`.

Package commands account for locked content before reporting that they apply. Wrap third-party or application commands with `canRunCommand(command, state)` when they also need to respect document locks.

Use `insertTab` to insert the same editable DOCX tab as the built-in Tab binding. It applies inside an ordinary paragraph; table cells and list paragraphs keep their navigation and level-changing Tab behavior.

## Tables

Table-cell controls use `canSetCellFormatting`, `activeCellVerticalAlign`, `setCellVerticalAlign`, `activeCellPadding`, and `setCellPadding`. The gate query covers the whole selection, alignment reports `top`, `center`, `bottom`, or `mixed`, and each padding side independently reports a point value, `mixed`, or `null`. A padding command writes only the sides supplied. `activeCellBorderColor`, `canSetCellBorderColor`, and `setCellBorderColor` operate on effective visible borders; editing a border inherited from its table or table style creates a direct override only on the visible sides of the selected cells.

## Images, links, and annotations

Image controls can use `IMAGE_FILE_ACCEPT`, `imageFilesIn`, and `insertImageFiles`. Lower-level controls can combine `readImageFile`, `fittedExtent`, and `insertImage`.

Link controls use `canSetLink`, `activeLink`, `activeLinkSpan`, `setLink`, and `removeLink`. External links are editable; bookmark links remain preserved.

Comment controls use `canAddComment`, `addComment`, `documentComments`, `selectComment`, `updateComment`, `removeComment`, `addCommentReply`, `updateCommentReply`, `removeCommentReply`, and `setCommentResolved`. `commentAuthor` sets the identity written by the built-in comment and reply composers; when omitted, the built-in UI writes `Anonymous`.

`documentNotes` returns the distinct footnotes and endnotes referenced by the main document, including their display labels and plain-text bodies.

## Presets and plugins

The root entry exports the preset values used by the built-in controls, including `DEFAULT_COLORS`, `DEFAULT_FONT_SIZES`, `DEFAULT_LINE_SPACINGS`, `DEFAULT_CELL_BORDERS`, and `DEFAULT_FONTS`. Pass replacement lists through the `presets` prop when the built-in controls should offer different values. Color presets remain quick choices; the built-in color controls also accept any three- or six-digit HEX color.

Pass ProseMirror plugins through the `plugins` prop. Consumer keymaps receive keyboard events before the built-in keymap, so an application can override bindings such as `Mod-k` without replacing the editor.

See [Feature support](./features.md) for behavioral limits and the published TypeScript declarations for complete signatures.
