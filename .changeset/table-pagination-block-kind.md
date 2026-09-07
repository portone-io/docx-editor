---
"@portone/docx-editor": patch
---

Table pagination moves into a block-kind module; no visible change. What the page engine knows about a table - where it may be parted between rows, the spacer and repeated header a continued page is drawn with, and which positions a page cut may stand at - used to be spread over the measurer, the decorations and the plugin state. It is now one module beside the paragraph's, and the engine asks whichever kind claims a block. Pages, page breaks and continued tables look exactly as they did.
