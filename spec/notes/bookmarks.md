# Bookmarks

## Range pairing

`w:bookmarkStart` and `w:bookmarkEnd` identify one bookmark by a matching required `w:id`. The start must occur before its matching end in document order. A missing matching marker makes the document non-conformant.

The bookmark name belongs to `w:bookmarkStart`. Where duplicate names occur, the first start marker in document order is maintained and later bookmarks with that name are ignored by a consumer.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§17.13.6.1–17.13.6.2.

## What we preserve

Bookmarks are cross-structure annotations and can span paragraphs. Markers inside a paragraph remain inline preservation nodes carrying `display: "hidden"`, so nothing of them is drawn. Markers directly under `w:body` are invisible block preservation nodes, so they retain their exact XML and order without becoming unsupported-content placeholders.

Both carry `guarded: true`, which is what the deletion guard answers for; [Preservation tiers](./preservationTiers.md) is where that attribute is decided.

Markers under `w:tbl` and `w:tr` have no node of their own, because those two levels hold only rows and cells. They ride along on the child before them - a row's markers on the cell they followed, a table's on the row - and are written back in the same spot, so a bookmark spanning a column no longer costs the table its structure. A marker following a cell that only continues a vertical merge is the exception: that cell is created fresh on export and has nothing to carry the marker, so the table is preserved whole instead.

Markers inside a container this reader could not take apart otherwise remain with that container's preserved XML. They are not exposed as independent model nodes.

A placeholder carries the identity of the session it was read from, and an export refuses one from another session as `lost-original`.

The editor does not create, delete, rename, or navigate to bookmarks. It keeps existing markers at their document-model positions while surrounding supported paragraphs are edited.

Table bookmarks can use the paired `w:colFirst` and `w:colLast` attributes on the start marker. These and `w:displacedByCustomXml` are retained in original XML rather than interpreted.
