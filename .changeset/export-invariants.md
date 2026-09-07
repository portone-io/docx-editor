---
"@portone/docx-editor": minor
---

Ask whether a document can be exported before trying.

`exportProblems(doc, session)` on the core entry, and `canExport(state)` with `documentExportProblems(state)` on the commands entry, report known reasons the writer would refuse the document, in the order it would raise them: each under the code and with the message the `DocxExportError` would carry, and with its position where the problem stands in the document. `exportDocx` throws the first entry of the same list, so problems reported by the query are also refused by the write.

The editor's handle answers the same list as `exportProblems()` beside `exportBytes()`, and `downloadDocx` asks it first and returns `{ status: "blocked", problems }` instead of throwing. A refusal the list does not foresee, such as a node a plugin stripped of an attribute the writer needs, is still thrown. `DownloadDocxResult` gaining a fourth member is a compile error for a `switch` over `status` written to be exhaustive; add a `blocked` branch that shows the problems.

A document whose comments part arrived as an empty element used to refuse its first comment with `malformed-xml`, since the writer looked for a closing tag the element does not have; the part is now opened for the entry, the way an empty extended comments part already was.
