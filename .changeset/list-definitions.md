---
"@portone/docx-editor": minor
---

A list that takes its numbers from a numbering style now draws them. Such a list holds no numbers of its own: it names a style (`w:numStyleLink`), the style names a list, and that list's definition is where the numbers are. Nothing followed that trail, so a list written this way - which is how Word writes a list built from its gallery - drew no marker at all.

More number formats are spelled out. Alongside decimal, bullets, upper and lower letters and lower Roman numerals, a list may now count in upper Roman numerals, decimal numbers with a leading zero, Ganada, Korean digits and the Chinese counting system. A format past those is still shown with decimal numbers.

Each level is drawn the way it asks to be: where its counting starts over (`w:lvlRestart`), whether the numbers in its text are all spelled as decimals (`w:isLgl`), what stands between its number and the paragraph text (`w:suff`), and where the number sits in the room kept for it (`w:lvlJc`).

A marker is also drawn in the character formatting its own level writes down (`lvl/rPr`): bold, italic, colour, size and typeface reach the number and never the text of the paragraph it stands in front of.

Everything here is read. The numbering part goes back out exactly as it arrived.
