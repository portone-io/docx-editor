---
"@portone/docx-editor": patch
---

Pasting from Word, Google Docs or LibreOffice now keeps tables as tables and Word lists as lists.

A table on the clipboard used to arrive as a single paragraph holding the text of every cell run
together. It now arrives as a table: one paragraph for each block a cell held, the columns and rows
a cell reached across kept, and the same width and lines a table inserted here has.

Word writes list items as ordinary paragraphs that name the list they belong to and draw their own
bullet or number, so pasting one used to give unnumbered paragraphs each beginning with a stray
marker. The items of one Word list now join one list, at the level Word gives them, counted when
the marker Word drew counts.

A table's caption is kept as the paragraph above it, a table inside a list item is read as a table
rather than as the item's own text, and a table too large for the editor to hold is read as one
paragraph per row instead of as a single run-on one.

A copy wrapped in a single element, which is how Google Docs writes one, no longer has its
paragraphs pressed into one.
