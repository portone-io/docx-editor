---
"@portone/docx-editor": patch
---

Enforce the lock a block-level content control carries, and refuse edits inside a group control.

A `w:sdt` standing under the body or inside a table cell was drawn as locked when its `w:lock` said so, and was then editable and deletable anyway. Its lock is now honored the way a wrapped cell's already was: where the control locks its contents, typing, formatting and deleting inside it are refused, a partial deletion that crosses its edge is refused with them, and so is an edit made in a table the control holds. A control locked against deletion is kept even when the selection covers it from end to end, while a `contentLocked` control may still be deleted whole and no less than whole. Unlocking from the editor now reaches a block control too, leaving the control itself standing under the id the file gave it.

A control that names `w:group` refuses edits to its contents even though it states no lock, which is what the specification asks of a group (§17.5.2.17). Its own paragraphs are refused, while a control standing inside it - a text control, a cell, or another block-level control - stays editable, and the group itself may still be deleted whole. There is no lock on such a control to lift, so the editor offers none, and `selectionLock` answers the new `shut` for a selection standing there rather than `none`.

A lock reaches every control standing inside the one that carries it, so text inside an open control held by a locked cell, by a locked block control or by a locked text control is refused as well.
