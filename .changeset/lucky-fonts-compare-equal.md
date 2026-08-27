---
"@portone/docx-editor": patch
---

Read two Unicode spellings of one font name as the same font.

A name a document writes down can be composed or decomposed (NFC or NFD) while the same name elsewhere in the document, or among the toolbar's presets, is written the other way.
Those spellings used to compare unequal, so the font dropdown listed the same font twice and a selection that all carried one font could report as mixed and blank the dropdown.
Names are now compared in composed form, and the spelling the document wrote is what stays stored, exported, and shown.

A list marker cut to its length cap is also cut between characters now, instead of possibly splitting an emoji or a combining sequence in half and drawing a replacement glyph.
