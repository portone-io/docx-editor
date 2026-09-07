---
"@portone/docx-editor": patch
---

A relationships part whose root carries a namespace prefix, or that arrived as an empty element, is now spliced correctly when the export adds a relationship to it: the entry goes inside the root, spelled under the root's own prefix, where the export used to refuse the file with `malformed-xml` for want of a bare closing tag. A numbering part that arrived as an empty element takes a new list definition the same way instead of being refused.

Every part the export rewrites is now read back as XML before the file is repacked, so a part that would not open is refused with `malformed-xml` naming the part rather than handed back.
