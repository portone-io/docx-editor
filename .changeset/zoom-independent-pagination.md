---
"@portone/docx-editor": patch
---

A document now breaks into the same pages however it is being looked at. Below a zoom of about 0.8 a document that sat close to a page boundary gained a page - in the demo, a blank one - and its pages moved as the window narrowed, because the editor scaled the paper in a way that measured a table taller at one zoom than at another. A zoom the application puts around the editor, such as a dialog that opens on a scale, no longer moves them either.
