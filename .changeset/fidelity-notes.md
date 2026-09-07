---
"@portone/docx-editor": minor
---

Read what a document holds that this editor cannot model. `importDocx` returns a `notes` array alongside the document and the session, the new `exportDocxReport` returns the same list beside the bytes it writes, and the new `documentFidelity(state)` on `./commands` answers the question about the document standing in an editor. A `FidelityNote` says how much of the original survived, what kind of content it was, the part and body block it came from, where in the document it stands, and the original element name, so a host can list what a file lost instead of guessing at it. Destructuring `importDocx` is unaffected.

A table now carries its `w:tblGridChange` instead of losing it whenever the table is rebuilt. The grid is still written from the column widths, and the revision markup closes it where CT_TblGrid takes it. `onlyCommentsChangedBy` therefore catches a submission that lost a `w:tblGridChange`, where it used to accept one: a returned file that dropped it now answers `body-changed` rather than `ok`.
