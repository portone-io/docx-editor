---
"@portone/docx-editor": patch
---

A paragraph copied inside the editor no longer exports a duplicate `w14:paraId`. The copy goes out as a paragraph of its own, without the original's identifiers, and the original still goes out as the bytes it arrived as.

A preserved block copied twice, a body-level bookmark marker or a section break among them, is refused with `unsupported-content` instead of written twice.
