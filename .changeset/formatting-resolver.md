---
"@portone/docx-editor": minor
---

Applying a paragraph style keeps the spacing the document defaults lay down, and formatting a style or the document defaults switch on can be switched off: the off is written into the run (`w:b w:val="0"`) and drawn as off, where it used to be dropped and the style's value drawn again.

Every display value a paragraph or a run carries is now resolved in the ECMA-376 §17.7.2 order by one resolver, whichever path built the paragraph: opening the document, applying a style, a paragraph edit, a paste, or typing into a fresh paragraph. A character style a run points at (`w:rStyle`) takes its place in that order, a hanging indent's implicit tab stop follows the indent actually drawn, and a document's `w:noTabHangInd` setting switches that stop off.

`RunFormat` grows to say so: `bold`, `italic`, `strike` and `smallCaps` are `boolean` (false is a toggle the run switches off outright), and `underline` may be `"none"`. A reader checking `=== true` or a truthy value is unaffected.

Supported character defaults now appear in the text and toolbar, so a default bold setting turns off with one press. Clearing a direct font size immediately restores the inherited paragraph or character style in the formatting controls.
