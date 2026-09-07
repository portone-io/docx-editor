---
"@portone/docx-editor": patch
---

Every node and mark attr is now classified by where its value comes from, and the documentation draws the line a plugin may rely on from that one classification rather than from prose. Node and mark names, `srcId`, the `format` family and the `to*Format` readers, and the exported types are the part that holds; the attrs carrying original OOXML - `pPr`, `rPr`, `tcPr`, `xml`, the `*Prefix` pair, `commentXml` - are this editor's bookkeeping, may change between releases, and can make an export refuse the node they sit on when their value is made by hand rather than read from a file.

No behavior changes.
