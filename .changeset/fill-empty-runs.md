---
"@portone/docx-editor": patch
---

Show and fill a highlighted run that holds no text, the blank a form leaves to be filled in.

A run with formatting and no characters, such as a yellow-highlighted empty table cell, used to be kept as invisible XML: nothing showed where the blank was, and text typed there came out unformatted.
Where the run's own properties paint a highlight or a shading, the editor now draws it as a short box, whose width the `--docx-editor-empty-run-width` property sets.
Any other empty run still takes no room, including one whose highlight or shading comes only from a style.

Clicking the box and typing, composing, or pasting fills it: the text takes the run's formatting, whatever formatting pasted text carried, and the box goes in the same undo step.
Bookmarks and other markers that draw nothing beside the box do not stop it filling, and a blank of other formatting beside it stays.
In a content control holding only such a run, the text goes into the control and what is typed next stays in it.
A blank inside a hyperlink, or in a control whose lock keeps its contents from being edited, is not filled; the text goes beside it in the formatting of the text around it.

A blank nobody filled goes back out in the file byte for byte as it came, even when the paragraph around it was edited.

Typed text now takes the formatting of the nearest text in its paragraph, the text before the caret or else the text after it, passing over what draws nothing, as Word does.
Text typed beside a bookmark, a comment's range marker, or a formatted run with no text used to take that element's formatting, which usually meant none.
In a paragraph with no text, typed text takes the properties of a formatted run with no text in it, including properties nothing draws, such as `w:rtl w:val="0"` that Google Docs writes, and the run's revision ids; such text is no longer drawn with the paragraph style's character formatting alone.
