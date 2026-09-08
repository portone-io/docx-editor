---
"@portone/docx-editor": patch
---

A new list is exported with the definition it was created with instead of one derived from its numbering id.

Starting a list used to leave nothing behind but the number the paragraph was given, and the shape of the list - numbered or bulleted, and the symbols each level counts in - was worked out again from whether that number happened to be even or odd. The number is now nothing but a name: the definition is recorded on the document when the list is started, drawn from as the markers are drawn, and written into numbering.xml as it stands. A list pasted from outside is defined the same way.

Undo takes a list's definition back with the edit that made it, and leaving a list gives back both the number and the definition it took, so a number is never spent on a list that is no longer there.

A document whose paragraph names a list that nothing defines - neither the file nor anything started while editing - is refused rather than written out with the list missing. `exportProblems` reports it ahead of the write, under `unsupported-content` and at the paragraph it stands at, the code that already covers a document holding what no correct file can be written from.

For `@portone/docx-editor/core`: `Numbering` now carries `added`, the definition of each list started while editing, beside the `lists` the file defines. Those definitions have the new `NewList` shape, and the levels of a definition are handed out as read-only maps.
