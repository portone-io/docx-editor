---
"@portone/docx-editor": patch
---

Stop text copied out of the editor from carrying the document's private data.

A copy was drawn the way the editor draws itself, so the HTML it left on the clipboard held the paragraph and run XML, the name and recorded identity of a comment's author, what the comment and its replies say, and the body of a footnote.
All of it landed in whatever application it was pasted into, and in `comment` mode a reader who may not change the body could take the body out this way.
A copy now carries only what a reader of it needs, and the paragraph style it names is the style's id rather than the whole `w:pPr`.

The plain text beside it says what was copied: a tab is a tab, a line break is a line, a page break is a form feed, and a table's cells stand apart by a tab and its rows by a line, so a table pasted into a spreadsheet arrives as a table.
