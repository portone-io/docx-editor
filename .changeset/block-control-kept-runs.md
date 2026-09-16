---
"@portone/docx-editor": patch
---

Let a kept run inside a block-level content control go when no page can hold it.

Paragraphs inside a control that ask to be kept with the next one (`w:keepNext`) closed the boundaries between them for good, so a kept run longer than a page had nowhere to part and the control ran off the sheet. Such a boundary is now closed only while a page could hold what is kept together, and let go whole where none can, which is how a run of kept paragraphs outside a control has always been treated.
