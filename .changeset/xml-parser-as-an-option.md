---
"@portone/docx-editor": minor
---

Hand `importDocx`, `exportDocx`, `onlyCommentsChangedBy`, `documentNumbering` and `parseNumbering` an `xmlParser` to read a document on a runtime that has no `DOMParser` global, instead of installing one.

A call given neither is refused with `DocxImportError` and the new import code `no-xml-parser`. It used to fail with a bare `ReferenceError`, which a server checking a file a counterparty returned could not tell apart from a document that arrived damaged. Reading a document no longer asks for a `Node` global at all, so `DOMParser`, however it is supplied, is the only thing the core entry needs from a DOM.

Each entry point settles its parser as the call comes in, so a runtime holding none is turned down before the bytes are looked at: bytes that are not a docx opened without a parser now report `no-xml-parser` where they reported `not-a-docx`. A parser that answers markup it cannot read by throwing, rather than by handing back a document holding a `parsererror`, is read as `malformed-xml` instead of having its own exception reach the caller.
