---
"@portone/docx-editor": patch
---

Copying and pasting inside the same document now keeps what it was copied as: paragraph and character formatting, tabs, breaks, images, links, and whole tables, instead of the part an HTML reader could make out. Copying a grid of cells and pasting it over selected cells fills them in place. Comment markers, bookmarks, and note references are not copied, so cutting and pasting them back drops them; a pasted list whose numbering this document does not define is given a number that it does. Content copied from another document, including a second editor on the same page, still comes in through the reader that reads the markup.
