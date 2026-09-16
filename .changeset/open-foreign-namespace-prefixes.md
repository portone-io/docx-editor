---
"@portone/docx-editor": patch
---

Open a document whose WordprocessingML is written under another prefix.

A file whose main part spells the wordprocessing namespace `ns0:` or leaves
it as the default namespace - what Python's `xml.etree` writes, and what
Word and Google Docs both open - used to be refused as content the editor
cannot write back. Such a package now opens: each part is rewritten to the
prefixes the editor writes, tag and attribute names only, before anything is
read, so a comment, a footnote, a header and the body all come through and
the file goes back out spelled `w:` throughout. A part that had already
bound `w`, `r` or another prefix the editor writes to a namespace of its own
is left exactly as it arrived, since respelling it would change what its
preserved markup means; the checks that applied before still decide such a
file, so a main part that binds `w` itself to some other namespace - the one
arrangement no respelling can reach - is refused as it was. A package
already spelling those prefixes is untouched and still exports byte for
byte.
