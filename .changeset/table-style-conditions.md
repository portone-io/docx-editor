---
"@portone/docx-editor": minor
---

A table style is now displayed part by part: the header row, the closing row, the first and last column, the corners and the banded rows take the shading, the lines and the text formatting the style dresses them with (`w:tblStylePr`), and `w:tblLook` decides which of those parts a table takes. The bands are as many rows or columns wide as the style says (`w:tblStyleRowBandSize`, `w:tblStyleColBandSize`), and the header row and the closing row are not banded with the rest.


Existing cell text follows its new table position after row edits. Character-format commands account for table-style inheritance, and replacing the formatting context refreshes both cell and text display values. Original conditional-formatting XML remains preserved.
