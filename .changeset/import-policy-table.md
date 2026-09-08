---
"@portone/docx-editor": minor
---

Keep markup the editor cannot model where it stood, instead of standing the paragraph around it down.

Opening a document used to walk a handful of element names and give up on the rest, and giving up
travelled outward until a whole paragraph became an "Unsupported content" placeholder. A word
processor writes a `w:lastRenderedPageBreak` on every page it lays out, so a real document opened
with one locked paragraph per page.

Every level of the document now has a preservation rule of its own. A run child stays inside its
run, a paragraph child beside the runs, a block stays a block, and only a marker or a wrapper
standing between the rows or the cells of a table - where there is no node to keep it in - still
stands the table down. What is left shows as a small box naming the element it stands for, or as
nothing where the file drew nothing, and the paragraph around it stays editable.

- A field code and its drawn result, a tracked insertion or deletion, a symbol, a positional tab
  and a content control the editor could not take apart are visible where they stand.
- Bookmarks, permission ranges, the markers of a moved passage and the pieces of a field cannot be
  removed by an edit; a tracked-change container, a simple field and a symbol can be selected and
  deleted whole.
- `documentFidelity` and the `notes` of `importDocx` now report `preserved-run-content`, and no
  longer report `paragraph-demoted` for a document being opened.
- New node `rawRunContent`, and `rawInline` gains the attrs `element`, `display`, `text` and
  `guarded`. A plugin reading the document model sees both.
- A footnote body, a comment body and a header now read a `w:cr` and a `w:noBreakHyphen` the way
  the document body does.
