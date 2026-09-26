---
"@portone/docx-editor": minor
---

Tell the application when an edit is refused, and let it theme locked content.

`DocxEditor` takes `onEditRefused`, called when typing, deleting, pasting,
dropping or formatting in the document body is turned down. It is handed an
`EditRefusal`: the rule that refused the edit (`lock`, `protection`,
`controlEdge`, `preserved` or `section`), what the edit would have done, and,
for a lock, the controls it reached with their `tag`, `alias`, `id`, lock and
level. Before this a refused keystroke did nothing at all and said nothing.

The right-click menus now say "Locked content can't be edited." when the
selection, or the table under it, holds locked content, which is why their
editing entries stand disabled.

Four custom properties theme how controls are drawn:
`--docx-editor-locked-background` and `--docx-editor-locked-outline` for locked
content, which keeps its yellow tint and orange outline by default, and
`--docx-editor-control-background` and `--docx-editor-control-outline` for a
content control that stays editable, which is not marked by default.

The yellow now shows over a locked table cell or row whose shading the
document states as `auto`, which used to hide it, and a cell's own shading
shows through the tint rather than hiding it.
