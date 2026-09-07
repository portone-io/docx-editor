---
"@portone/docx-editor": patch
---

Enter no longer duplicates paragraph identifiers or a paragraph-level section break. The continuing paragraph keeps its identifiers, and the section break stays on the last paragraph of the split.

Deleting or joining away a section-ending paragraph is refused until section editing is supported. Replacing text within that paragraph remains available.
