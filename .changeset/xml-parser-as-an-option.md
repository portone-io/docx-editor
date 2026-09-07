---
"@portone/docx-editor": minor
---

Hand `importDocx`, `exportDocx` and `onlyCommentsChangedBy` an `xmlParser` to read a document on a runtime that has no `DOMParser` global, instead of installing one.

A call given neither is refused with `DocxImportError` and the new import code `no-xml-parser`. It used to fail with a bare `ReferenceError`, which a server checking a file a counterparty returned could not tell apart from a document that arrived damaged. Reading a document no longer asks for a `Node` global at all, so `DOMParser`, however it is supplied, is the only thing the core entry needs from a DOM.
