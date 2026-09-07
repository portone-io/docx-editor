---
"@portone/docx-editor": minor
---

A paragraph marked keep with next (`w:keepNext`), by its own properties or by its style, stays on the same page as the start of the block after it in the page guides, as it does in Word. A run of such paragraphs moves together with the first piece of the block the keeps end at; a run no page can hold is laid out as if no keep were set. The document is not changed: the mark is read, never written.

`ParagraphFormat` grows `keepNext?: true` to say so, beside `pageBreakBefore`.
