---
"@portone/docx-editor": patch
---

Hold the edge of a block-level content control, and act on the two properties that say a control does not outlive an edit.

Backspace at the start of a control's first block and Delete at the end of its last one no longer carry a paragraph out of the control or pull the block after it in; they do nothing, and the caret stays where it was. A selection running from outside a control into it, or the other way round, is refused rather than closed by moving text across the boundary. A control left holding a single empty paragraph is removed whole instead, unless its lock says it may not be deleted.

A control carrying `w:temporary` now loses its wrapper on the first edit inside it, as the specification requires, and one carrying `w:showingPlcHdr` stops claiming that what it holds is placeholder text, so a document exported after typing into it no longer shows that text as a placeholder in Word. Both go into the history with the edit that caused them, so one undo takes them back together, and a comment left on a document changes neither.

Placeholder text itself is not selected or replaced when a control stops claiming it holds a placeholder, so it is edited and deleted like any other text.
