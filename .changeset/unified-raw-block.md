---
"@portone/docx-editor": minor
---

Keep a table whose markers stand between its rows or its cells, and stand every block the editor
cannot model as one kind of placeholder.

**Breaking:** the `docxRaw` and `bookmarkBlock` node types no longer exist. Both were placeholders
for a block nobody could read, differing only in where the block stood and whether anything of it
was drawn, and `rawBlock` is now the one node for all of them. A plugin that matched either name
matches `rawBlock` instead and reads `display` (`"chip"` for a placeholder box, `"hidden"` for a
marker that draws nothing) and `guarded` to tell them apart. The class names
`docx-editor-bookmark-block`, `docx-editor-raw-xml` and `docx-editor-table` are gone with them;
every placeholder now draws as `docx-editor-raw-block`.

- A bookmark spanning a table column stands under the row rather than inside a cell, and used to
  cost the whole table its structure. Such a table now opens as a table, and the marker goes back
  exactly where it stood. A marker following a cell that only continues a vertical merge is the
  remaining exception, since that cell is created fresh on export.
- `documentFidelity` and the `notes` of `importDocx` no longer report `table-demoted` for a table
  whose only unread markup is a marker.
- A placeholder can now be moved between the body and a table cell and still export.
- `table` and `tableRow` gain a `leadingXml` attr, `tableRow` and `tableCell` a `trailingXml` attr:
  the markers each carries between its children.
