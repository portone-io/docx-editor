---
"@portone/docx-editor": patch
---

Content control properties are written in the order the schema lays down. Locking a control that declares what kind of control it is - a date picker, a drop-down, plain text - used to write `w:lock` after that declaration, where CT_SdtPr puts `w:lock` before it, so a validator reading the exported file could refuse a control this editor had rewritten. A control carrying `w:label` or `w:tabIndex` was written the same wrong way round.

Nothing else about an exported file moves: a document that goes out untouched is still byte for byte the one that came in, and every other edited fragment is written exactly as before.
