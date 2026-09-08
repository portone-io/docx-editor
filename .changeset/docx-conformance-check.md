---
"@portone/docx-editor": minor
---

ECMA-376 Strict packages are refused with the new `unsupported-conformance` code, and a document whose main part root does not bind `w` to Transitional WordprocessingML is refused as `unsupported-content`. Both used to open far enough to fail later and obscurely: a Strict package as `missing-part`, and a document under another prefix as a `malformed-xml` about bookmarks on its first edit. Word saves neither by default. The comment and hyperlink writers now have the part's root declare the prefix they write under, rather than declaring it on every element written, so an untouched document still exports byte for byte.

A write requiring a prefix already bound to a different namespace is refused instead of silently giving new markup the wrong meaning. New hyperlink relationship attributes are also checked for declarations that shadow the root binding.

Editing existing numbering and comment extension parts that use another prefix now adds the writer's required declaration at the root, so the exported XML remains readable.
