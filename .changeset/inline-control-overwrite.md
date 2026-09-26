---
"@portone/docx-editor": patch
---

Keep an inline content control when everything it holds is written over or deleted.

Selecting all of the text inside a content control that stands in a paragraph - a bracketed placeholder a template leaves to be filled in, say - and typing over it used to take the control away: the new text landed beside it as ordinary text, and the tag and id a server looks for when it fills the document in were gone. The same happened with a drag carried past the end of the line, a Shift-click, a triple click, a composition, and a paste. The new text now goes into the control, which keeps its tag, its id, and everything else it states, and what is typed next goes on into it.

Deleting everything the control holds now leaves the control in the file with nothing inside it, and text typed right after goes back into it. A control locked against deletion alone (`sdtLocked`) takes all of this, since its contents may be edited; before, deleting or retyping everything it held was refused. A selection that also takes in text outside the control still removes the control with that text, and a control locked against deletion still refuses it.
