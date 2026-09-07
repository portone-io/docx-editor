---
"@portone/docx-editor": patch
---

Universal measures such as `8.5in` are read correctly. A measurement in a document may be written as a length with a unit - `8.5in`, `2.54cm`, `12pt` - as well as a count, and every one of them used to be read as its leading digits alone. A US Letter document whose section says `w:pgSz w:w="8.5in"` was drawn as a page 8.5 twips wide, which is no page at all, so the editor fell back to A4 and showed the wrong paper; a tab stop at `1.5in` landed at 0.08pt and an automatic tab interval of `0.75in` collapsed every tab in that document to no width. Page size and margins, indents, spacing, font sizes, table and cell widths, cell margins, row heights and tab stops are now read as the lengths they name.

`on` and `off` spellings of boolean attributes are read as the schema admits. A style marked `w:default="on"` is now recognised as the default style for its kind, as `w:default="1"` already was, so a document that marks its defaults that way is shown with the formatting they lay down.

Integer measurements retain support for an explicit `+` sign. Table and cell widths with an explicit `%` follow Word's percentage interpretation even when their width type says otherwise. Rebuilt table widths and grid columns are written in whole units; untouched XML is preserved.
