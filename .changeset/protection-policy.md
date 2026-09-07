---
"@portone/docx-editor": patch
---

`onlyCommentsChangedBy` reads which parts a comment edit may rewrite, what an entry in one of them may look like, and how the document story is compared off a single policy object, and that object is the one the editor's own part planners write those parts from. The verdicts are the same verdicts and `CommentOnlyVerdict` is unchanged.

Two shapes it used to answer `ok` for now come back refused, both of them markup this editor never writes. Bytes riding between the entries of a comment part - a comment, a stray run of text, a CDATA section - are read as what they are rather than passed over; whitespace still is not, so a part laid out over several lines comes back fine. And a namespace declaration is judged by what it binds a prefix to, on an entry and on the part's own root element alike, so a submission cannot rebind `w14` under this editor's own markup.
