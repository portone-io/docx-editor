---
"@portone/docx-editor": patch
---

Give a document's sections a model of their own. Every `w:sectPr` is now read once - its paper and
margins, the header and footer stories it selects, whether the first page differs, and the page
number it starts at - instead of being read twice by two readers that looked for it differently.

The section that closes the body rides on the document node as `sectPr` rather than in the
preserved tail. It is written straight back out, so an untouched document still exports byte for
byte, and a submission that rewrites the paper it is written on is refused by
`onlyCommentsChangedBy` as it was before.

Pressing Enter in the last paragraph of a section continues to leave the break on the later half,
now on the same reading of `w:pPr` the rest of the package uses.
