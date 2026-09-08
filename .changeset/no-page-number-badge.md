---
"@portone/docx-editor": patch
---

Stop drawing the page number on the corner of each page.

The page guides laid a small grey number inside the top right corner of every page, over the paper the text sits on.
It is gone, and the guides now draw the gaps between pages and the header and footer stories alone.
A document that prints its own page number through a `PAGE` field in a header or footer is unaffected: that number is the document's, not the editor's.

The `docx-editor-page-badge` class the number carried is no longer emitted, so a rule of your own written against it in the published stylesheet no longer matches anything.
