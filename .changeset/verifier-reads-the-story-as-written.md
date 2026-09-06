---
"@portone/docx-editor": patch
---

Stop `onlyCommentsChangedBy` from refusing a comment written into a table cell.

A commented table is rebuilt on the way out, so it came back worded the way this editor words it while the original carried the wording its producer chose, and the two were compared word for word.
The story is now compared as this editor writes it back, so two blocks this editor would write alike are read alike: the attribute order inside a `w:tcW` or a `w:tblW`, a percentage width written as `100%` or as `5000`, runs a producer split that say the same text, and a table's `w:tblGridChange`.
Formatting properties keep the line breaks, comments and text a producer wrote between them, so rewriting a run's bold, a paragraph's alignment or a cell of a rebuilt table no longer drops them.
A rebuilt table keeps what stood inside its width and span properties too.
