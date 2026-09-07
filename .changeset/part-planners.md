---
"@portone/docx-editor": patch
---

A relationships part whose root carries a namespace prefix, or that arrived as an empty element, is now spliced correctly when the export adds a relationship to it: the entry goes inside the root, spelled under the root's own prefix, where the export used to refuse the file with `malformed-xml` for want of a bare closing tag. A numbering part that arrived as an empty element takes a new list definition the same way instead of being refused.

When one export adds several parts, `[Content_Types].xml` now declares them in the order they were added, a media type's `Default` ahead of any `Override`; each writer used to put its own declaration first, so the declarations came out in the reverse order of adding. Every declaration is still placed right after the opening tag, and the rest of the part is left as it arrived.

Every XML part the export rewrites is now read back before the file is repacked, so a part that would not open is refused with `malformed-xml` naming the part rather than handed back.

New part names and content-type requests recognize names differing only in case, avoiding duplicate package entries and declarations.
