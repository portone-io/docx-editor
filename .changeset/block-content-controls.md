---
"@portone/docx-editor": patch
---

Open a block-level content control as the blocks it holds.

A `w:sdt` standing under the body or inside a table cell used to open as a placeholder: a box saying the content was preserved, with the paragraphs and tables inside it neither readable on the page nor editable. It now opens as the content it wraps, so the text of a clause held in a rich text control reads and edits like any other text, and a control holding a table holds a table. The control's own opening XML - its id, its lock, its type, its data binding - is kept whole and written back untouched, and a control nobody edited still goes out as the bytes it arrived as.

Controls nest as the file nested them, so a `w:group` around a rich text control comes back that way round, and a control inside a table cell keeps the formatting the cell dressed it with. A control that is duplicated gets an id of its own on the way out, and its data binding stays with the first copy, so two controls never edit each other.

Two kinds are still preserved whole, because opening them would let an edit write a file Word reads differently: a `w:text` control, whose content may be no more than one paragraph, and a `w:picture` control, whose content may be no more than one picture.

Locks are read off a block control but not yet enforced against an edit, and a control does not yet break across a page.
