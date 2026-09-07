---
"@portone/docx-editor": patch
---

An edit to a cell border, a cell shading, a row height, a paragraph indent, or a line spacing now reads and writes the WordprocessingML attribute alone. A producer's own attribute that shares the local name (`x:val` beside `w:val`, declared ignorable) used to be taken for the formatting value, or written over in its place; it is now left as the producer wrote it, and the value Word reads is the one read and written.
