---
"@portone/docx-editor": patch
---

Formatting XML the editor writes now escapes every attribute value, so a style id holding `&` no longer produces an unreadable file; nothing else changes.
