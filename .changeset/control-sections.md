---
"@portone/docx-editor": patch
---

End a section at the paragraph inside a content control that carries its break, and read every break a control holds.

A section break written on a paragraph inside a block-level `w:sdt` now ends its section at that paragraph, as it does for a paragraph outside a control, and every break a control holds is read.
What the control holds after a break starts the next section's page, on that section's paper and under its header and footer, as Word parts the control across the next sheets, while the control stays one control in the document.
Footnote and endnote numbering that restarts per section, and the width a table is fitted to, follow the section the paragraph is in.
