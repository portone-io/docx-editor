---
"@portone/docx-editor": minor
---

`srcId` now names the block within the session it was opened in. A preserved block used to carry a bare index into the blocks of whichever document it was exported against, so a block moved or pasted in from another document pointed at this document's block of the same number and went out as that block's XML instead of its own. The attribute is now a string naming the document, the story and the place in it, and a block whose document is not the one being exported is refused with `lost-original` rather than written as something else.

A plugin that read `node.attrs.srcId` as a number needs to change: it is a string, and the only thing to do with it is hand it back as it was found. Nothing else about the attribute is public, and no exported type or function signature changed.
