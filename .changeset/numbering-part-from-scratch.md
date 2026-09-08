---
"@portone/docx-editor": minor
---

A new list no longer needs the document to have arrived with a numbering part. The export writes `word/numbering.xml`, relates it from the main part and declares it in the content types, so a document that never held a list can take one; the list commands, which used to report that they did not apply in such a document, now do. The `missing-numbering-part` error code is removed, as nothing can reach it any more.

A package that has no `[Content_Types].xml` at all still turns a new list down, since the new part has nowhere to be declared. That refusal comes back as `missing-content-types`, the code a new image and a new comment part are already refused with, and `exportProblems` reports it ahead of the write as before.
