---
"@portone/docx-editor": patch
---

Copying and pasting inside the same document now keeps what it was copied as: paragraph and character formatting, tabs, breaks, images, links, and whole tables, instead of the part an HTML reader could make out. Copying a grid of cells and pasting it over selected cells fills them in place. Comment markers, bookmarks, and note references are not copied, so cutting and pasting them back drops them. A pasted list whose numbering the document no longer defines is given a number that it does, keeping the bullets or numbers it was copied with; a list arriving from another document or application is renumbered by the reader that reads it. Content copied from another document, including a second editor on the same page, still comes in through that reader.
