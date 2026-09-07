---
"@portone/docx-editor": patch
---

Commands now share guard helpers so their applicability checks and dispatched edits respect the same rules. Formatting at a caret inside locked content is refused; formatting queries continue to report the selected text's values under document protection.
