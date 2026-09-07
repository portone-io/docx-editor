---
"@portone/docx-editor": minor
---

Run formatting is read, written and compared through one property table, so the value a control reads, the XML an edit writes and the check that leaves text already in that state alone can no longer drift apart. The XML written is byte for byte what it was.

`RunFormat` gains `caps`, `doubleStrike` and `characterSpacingPt`, read off a run's `w:caps`, `w:dstrike` and `w:spacing`. They reach the `format` attr a plugin reads; the editor does not draw or edit them yet.
