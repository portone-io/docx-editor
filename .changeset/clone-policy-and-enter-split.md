---
"@portone/docx-editor": patch
---

Enter no longer duplicates `w14:paraId` or a paragraph-level `w:sectPr`; a paragraph carrying a section break can no longer be joined away.

Splitting a paragraph handed the new one every attr of the original, so two paragraphs went out claiming the same `w14:paraId`, which is the name a comment anchors to, and a section break was written twice, giving the document a section it never had. The break now stays on whichever half ends the section, and the new paragraph goes out with no identifier of its own.

Joining that paragraph into the one above it - a Backspace at its start, or a selection run across its boundary - is refused, since nothing the editor writes can put a lost section back. A selection running across the last paragraph of a section therefore cannot be deleted at all for now, which [Features](/docs/features) records as a limitation.
