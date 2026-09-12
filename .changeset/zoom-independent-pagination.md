---
"@portone/docx-editor": patch
---

Break a document into the same pages whatever zoom it is read at. The paper is now scaled with a transform rather than the CSS `zoom` property, which laid every box out on whole device pixels of the scaled rendering and measured a table taller at one zoom than at another.
