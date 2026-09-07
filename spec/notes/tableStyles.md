# Table styles

A table style dresses the parts of a table separately: the header row, the closing row, the edge columns, the corners and the banded rows each have their own `w:tblStylePr` (§17.7.6.6), and `w:tblLook` on the table says which of those the table takes (§17.4.55). Which parts one cell belongs to is worked out in `src/docx/tableFormatting/conditions.ts`, and every caller that asks what a cell falls back on reads that one answer.

## The order the parts are applied in

§17.7.6 fixes the order, a later one overriding an earlier one: whole table, banded columns, banded rows, first row and last row, first column and last column, then the four corners. A cell in the header row of the first column therefore takes the header row's formatting over the column's, and the corner over both.

The table style's own `w:tblPr`, `w:pPr` and `w:rPr` lie under all of them, and the whole style sits where §17.7.2 puts it: above the document defaults and below the numbering level, the paragraph style and direct formatting. A paragraph style that states an alignment therefore beats one a conditional format states, which is what the hierarchy asks for even where Word's own drawing of a built-in style suggests otherwise.

## Which parts a table takes

`w:tblLook` states the six settings as attributes (§17.4.55). A Transitional document may write them instead as a hexadecimal bitmask in `w:val` (Part 4 §14.4.12): `0x0020` first row, `0x0040` last row, `0x0080` first column, `0x0100` last column, `0x0200` no row banding, `0x0400` no column banding. An attribute that is there answers for the bit of the same name; where it is absent the bitmask answers, and a `w:val` that is no bitmask at all, like the element being absent, is read as `0000`.

`0000` is not "no conditional formatting". Both settings that switch the banding off are worded as `no`, so all bits clear leaves the row and column banding applied and takes only the header row, the closing row and the edge columns away - which is also what §17.4.55 states as the default for a table writing no `w:tblLook`. A table that wants nothing at all writes `0600`.

## Counting the bands

`w:tblStyleRowBandSize` and `w:tblStyleColBandSize` (§17.7.6.5, §17.7.6.7) say how many rows or columns make up one band, and default to one. Band 1 is the first band, so band 1 and band 2 alternate from the first banded row onwards.

The specification does not say whether the header row and the closing row are counted among the banded rows. Word does not band them: with the header row taken, the first banded row is the one under it and it is band 1. This editor follows Word, for the rows and for the edge columns alike, because a header row banded as well would be drawn in the band's fill under Word's own built-in styles. Observed 2026-09-08 against ECMA-376 5th edition, Part 1 §§17.4.55, 17.7.6, and Part 4 §14.4.12.

## The lines a conditional format draws

The `w:tcPr` of a conditional format writes the lines of the part it dresses rather than of every cell in it: its four sides are drawn where a cell lies on the edge of that part, and `insideH` and `insideV` between the cells within it. That is why a header row whose conditional `insideV` is `nil` is drawn as one unbroken band, and it is the same rule the table's own `w:tblBorders` are drawn by. A `w:tblPr` inside the conditional format speaks for the same part and lies under its `w:tcPr`.

## `w:cnfStyle` is a record, not an instruction

`w:cnfStyle` on a row, a cell or a paragraph records, as a bitmask, which parts of the table the producer decided that object belonged to (Part 4 §14.4.9, §14.4.10). It is a cache of a decision the table's own grid and `w:tblLook` already determine, so this editor works the parts out from where the cell sits and never reads the record. The markup is preserved with the rest of the `w:tcPr`, so a document that carries one goes back out with it.

## What is not read

`w:tblStylePr` may also carry a `w:trPr`, and a row can carry `w:gridBefore` and `w:gridAfter` to stand short of the grid. Neither is read. Changing which style a table wears is not an edit this editor offers yet; the display is derived from what the document already says.
