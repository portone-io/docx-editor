---
"@portone/docx-editor": patch
---

Every section now previews the headers and footers it names. A document whose second section
selects a header of its own used to draw the first section's on every page, because only the first
section's references were ever resolved.

A header and a footer body is read the way the document body is: the same block readers, the same
verbatim slices, one story per part on the document node beside the comment and footnote stories.
The plain-text walker each of them used to have is gone, so a `w:cr`, a no-break hyphen and a field
read the same in a header as they do anywhere else, and a `PAGE` or `NUMPAGES` field is now found by
pairing the field characters it is written between rather than by scanning the text they surround.

A header story the editor rewrote is written back into its own part, around the two ends that part
arrived with, and every part nobody rewrote is repacked as the bytes it came as. Editing a header is
not offered in the editor yet; this is the model and the export path underneath it.
