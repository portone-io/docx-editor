---
"@portone/docx-editor": patch
---

Stop `onlyCommentsChangedBy` from refusing a comment written into a table cell.

A commented table is rebuilt on the way out, so it came back worded the way this editor words it while the original carried the wording its producer chose, and the two were compared word for word.
The story is now compared as this editor writes it back, so what the two files called the same thing is read as the same thing.
The cost is that markup this editor does not write back is not compared either: the attribute order inside a `w:tcW` or a `w:tblW`, a percentage width written as `100%` or as `5000`, and a table's `w:tblGridChange`.
