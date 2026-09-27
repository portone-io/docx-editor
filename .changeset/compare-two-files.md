---
"@portone/docx-editor": patch
---

`compareDocx` on the core entry lists the paragraphs, tables, comments and package parts that differ between two files.
`onlyCommentsChangedBy` now also refuses a file that declares a comment content type for anything but a comment part it holds, including by extension, or relates a comment part it does not hold.
