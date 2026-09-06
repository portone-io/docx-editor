---
"@portone/docx-editor": patch
---

Commands and `canRunCommand` now report false where a bookmark marker or a note reference would be removed, instead of reporting true and changing nothing.

A refusal over one of those markers also ends an open IME composition, the way a refusal over a locked control already did.
