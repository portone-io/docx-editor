---
"@portone/docx-editor": minor
---

Draw the table styles and list numbering a document defines, and let a document that never held a list take one.

A table style is now displayed part by part: the header row, the closing row, the first and last column, the corners and the banded rows take the shading, the lines and the text formatting the style dresses them with (`w:tblStylePr`), and `w:tblLook` decides which of those parts a table takes.
The bands are as many rows or columns wide as the style says (`w:tblStyleRowBandSize`, `w:tblStyleColBandSize`), and the header row and the closing row are not banded with the rest.
Text already in a cell follows its new position after a row edit, and the conditional-formatting markup a document arrived with goes back out untouched.

A list that takes its numbers from a numbering style now draws them.
Such a list holds no numbers of its own: it names a style (`w:numStyleLink`), the style names a list, and that list's definition is where the numbers are, which is how Word writes a list built from its gallery.
Alongside decimal, bullets, upper and lower letters and lower Roman numerals, a list may now count in upper Roman numerals, decimal numbers with a leading zero, Ganada, Korean digits and the Chinese counting system; a format past those is still shown with decimal numbers.
Each level is drawn the way it asks to be: where its counting starts over (`w:lvlRestart`), whether the numbers in its text are all spelled as decimals (`w:isLgl`), what stands between its number and the paragraph text (`w:suff`), and where the number sits in the room kept for it (`w:lvlJc`).
A marker is also drawn in the character formatting its own level writes down (`lvl/rPr`), so bold, italic, color, size and typeface reach the number and never the text of the paragraph it stands in front of.
All of this is read; the numbering part goes back out exactly as it arrived.

A new list no longer needs the document to have arrived with a numbering part.
The export writes `word/numbering.xml`, relates it from the main part and declares it in the content types, so a document that never held a list can take one; the list commands, which used to report that they did not apply in such a document, now do.
If an existing numbering relationship points to a missing part, the new definition is written at that location so the document can find the list when it is reopened.

A new or pasted list is exported with the definition it was started with, rather than a format inferred from its numbering IDs, and undo restores that registration along with the edit.
A list whose definition is missing or unsupported is refused as `unsupported-content`, and a package with no `[Content_Types].xml` for the new numbering part to be declared in is refused as `missing-content-types`; `exportProblems` reports either ahead of the write.
List and abstract definition IDs stay distinct in a document already using the largest safely representable integer, and a value that cannot be written faithfully is rejected rather than silently changed or discarded.

An ECMA-376 Strict package is refused with the new `unsupported-conformance` code, and a document whose main part root does not bind `w` to Transitional WordprocessingML is refused as `unsupported-content`.
Word saves neither by default, and both used to open far enough to fail later and obscurely: a Strict package as `missing-part`, and a document under another prefix as a `malformed-xml` about bookmarks on its first edit.
A write needing a prefix already bound to a different namespace is refused instead of silently giving new markup the wrong meaning, and a new hyperlink relationship attribute is checked for a declaration that shadows the root binding.
The comment and hyperlink writers now have the part's root declare the prefix they write under rather than declaring it on every element written, and editing an existing numbering or comment extension part that uses another prefix adds the writer's declaration at the root, so an untouched document still exports byte for byte and the exported XML stays readable.

`DocxImportErrorCode` gains `unsupported-conformance`.
`DocxExportErrorCode` loses `missing-numbering-part`, since nothing can reach it any more, and that is a compile error for a `switch` over the code written to be exhaustive; drop the branch.
`missing-content-types` now also covers the numbering part a new list is defined in.
On `./core`, `Numbering` gains `added`, and the new `NewList` and `NewListLevel` types describe the definitions the editor registered.
Level maps are read-only, and a registered definition carries a restart, legal numbering and a suffix, while marker run formatting and custom tab stops stay outside what can be registered.
