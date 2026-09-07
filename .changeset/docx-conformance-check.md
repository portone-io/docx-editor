---
"@portone/docx-editor": minor
---

ECMA-376 Strict packages are refused with the new `unsupported-conformance` code, and a document that binds WordprocessingML to a prefix other than `w` is refused as `unsupported-content`. Both used to open far enough to fail later and obscurely: a Strict package as `missing-part`, and a document under another prefix as a `malformed-xml` about bookmarks on its first edit. Word saves neither by default. The comment and hyperlink writers now have the part's root declare the prefix they write under, rather than declaring it on every element written, so an untouched document still exports byte for byte.
