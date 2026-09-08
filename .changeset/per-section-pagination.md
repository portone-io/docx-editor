---
"@portone/docx-editor": patch
---

Page boundaries and table widths now follow the paper of the section a block sits in. A document
whose second section is landscape used to be paginated as though the whole of it were the first
section's portrait paper, and a table put in that section was divided into the upright body width
and could not be widened past it.

Each page is now as tall as the body its own section leaves, a section break starts a new page
unless it is marked continuous, and a table is fitted to the body width of the section it is
inserted or dragged in. An image pasted into a section is fitted to the height that section's paper
leaves. A page number still carries on across a section boundary rather than restarting where the
section asks.

One sheet is still drawn at one width, the first section's, because nothing draws two paper widths
on one sheet. A landscape section is therefore paginated on landscape pages but drawn on the
sheet's width, and a table wider than the sheet's body is drawn shrunk to fit it while the document
keeps the width it was given. An even-page or odd-page start is preserved and exported untouched
but drawn as an ordinary new page.
