---
"@portone/docx-editor": patch
---

Content control properties are written in the order the schema lays down. Locking a control that declares what kind of control it is - a date picker, a drop-down, plain text - used to write `w:lock` after that declaration, where CT_SdtPr puts `w:lock` before it, so a validator reading the exported file could refuse a control this editor had rewritten. A control carrying `w:label` or `w:tabIndex` was written the same wrong way round.

Untouched documents retain their original XML.
