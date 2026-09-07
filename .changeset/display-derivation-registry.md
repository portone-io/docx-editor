---
"@portone/docx-editor": patch
---

The lines of a table's cells and the style values of a paragraph are worked out again by one plugin, which also works every value out again when the document's formatting is replaced under it; no behavior change for a document opened today.

Changing the formatting context also refreshes existing run marks, so text does not keep its previous style after the paragraph display values change. These updates preserve the original run XML, including in locked content and protected documents.
