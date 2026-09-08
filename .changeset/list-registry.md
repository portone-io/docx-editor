---
"@portone/docx-editor": minor
---

New and pasted lists are exported with their registered definitions instead of a format inferred from their numbering IDs. Undo restores the registration along with the edit. A new list with a missing or unsupported definition is refused as `unsupported-content` and reported by `exportProblems` before export.

List and abstract definition IDs remain distinct when a document already uses the largest safely representable integer. Values that cannot be written faithfully are rejected instead of silently changed or discarded.

On the core entry, `Numbering` gains `added` and the new `NewList` and `NewListLevel` types describe registered definitions. Level maps are read-only; new definitions support restart, legal numbering and suffix values, while marker run formatting and custom tab stops remain excluded from registration.
