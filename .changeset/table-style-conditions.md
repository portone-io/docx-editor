---
"@portone/docx-editor": minor
---

A table style is now displayed part by part: the header row, the closing row, the first and last column, the corners and the banded rows take the shading, the lines and the text formatting the style dresses them with (`w:tblStylePr`), and `w:tblLook` decides which of those parts a table takes. The bands are as many rows or columns wide as the style says (`w:tblStyleRowBandSize`, `w:tblStyleColBandSize`), and the header row and the closing row are not banded with the rest, the way Word draws them.

A table written by an older producer, which names the parts it takes through the legacy `w:tblLook w:val` bitmask instead of the six attributes, is read the same way.

The paragraphs inside a cell wear what the style lays down for the parts that cell belongs to, so text in a header row is drawn in the header row's formatting, whether the document arrived with it or an edit built the paragraph. Editing a cell measures against the lines those parts draw as well: coloring the border of a banded cell writes down the band's own line rather than doing nothing.

The document is not changed: everything here is read, and the `w:cnfStyle` a producer wrote on a cell is preserved rather than believed - which part of a table a cell belongs to is worked out from where it sits.
