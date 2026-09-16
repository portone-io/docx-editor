---
"@portone/docx-editor": patch
---

End a section at the paragraph inside a content control that carries its break, and read every break a control holds.

A section break written on a paragraph inside a block-level `w:sdt` used to close its section only at the end of the whole control, and a second break inside the same control was kept in the file without being read. The section now ends at the paragraph carrying the break, as it does for a paragraph outside a control, and each break a control holds lays down a section of its own. Footnote and endnote numbering that restarts per section therefore restarts where the file says, and a table standing after such a break inside the control is fitted to the width of the section it is in.

The control itself is still drawn on the paper of the section it starts in, and the next page still opens after the whole control.
