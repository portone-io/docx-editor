---
"@portone/docx-editor": patch
---

Copied HTML and text now follow one declared outbound shape per node and mark, so what leaves the editor is decided beside the schema rather than filtered out of the page's own drawing afterwards. A copy pasted into another application carries paragraphs, styles, runs, links, images at the size they were drawn, and tables as tables, and none of the document's internals. A comment marker, a bookmark, and the placeholder standing for content the editor could not model now leave nothing behind them, where before they left an element carrying the editor's own name for what stood there.

Copied plain text reads as the page does. A table pasted into a spreadsheet keeps one row per row, with everything a cell holds on the one line that cell stands on, and a block the editor only kept no longer opens a blank line where it stood. A kept line break and a kept character, such as a no-break hyphen, travel with the text they stand in, and so does the text a placeholder draws: a field's cached result and the words of a tracked insertion read in a copy as they read on the page.
