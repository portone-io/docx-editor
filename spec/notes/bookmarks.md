# Bookmarks

## Range pairing

`w:bookmarkStart` and `w:bookmarkEnd` identify one bookmark by a matching required `w:id`. The start must occur before its matching end in document order. A missing matching marker makes the document non-conformant.

The bookmark name belongs to `w:bookmarkStart`. Where duplicate names occur, the first start marker in document order is maintained and later bookmarks with that name are ignored by a consumer.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§17.13.6.1–17.13.6.2.

## What we preserve

Bookmarks are cross-structure annotations and can span paragraphs. Markers inside a paragraph remain invisible inline preservation nodes. Markers directly under `w:body` are invisible block preservation nodes, so they retain their exact XML and order without becoming unsupported-content placeholders.

Markers inside an unsupported container remain with that container's preserved XML. They are not exposed as independent model nodes.

The editor does not create, delete, rename, or navigate to bookmarks. It keeps existing markers at their document-model positions while surrounding supported paragraphs are edited.

Table bookmarks can use the paired `w:colFirst` and `w:colLast` attributes on the start marker. These and `w:displacedByCustomXml` are retained in original XML rather than interpreted.
