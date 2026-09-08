---
"@portone/docx-editor": minor
---

Comment bodies keep their formatting when edited, and `setCommentBody` accepts a formatted body.

**Breaking:** the `commentReference` node loses its `text`, `commentXml` and `imported` attrs and
the `noteReference` node loses its `text` attr, together with the `data-comment-text`,
`data-comment-xml`, `data-comment-imported` and `data-note-text` attributes that carried them into
the DOM. A plugin that matched any of those reads what a comment or a note says through
`documentComments` and `documentNotes` instead, which answer as they always did.

**Breaking:** `DocumentComment`, `DocumentCommentReply` and `DocumentNote` declare every field
`readonly`. They are the editor's own records, worked out once per edit and handed out rather than
copied, so a caller that was writing into one copies it first.

A side story - a comment's body, a footnote's - is now read the way the document body is: the same
block readers, the same verbatim slices, one document of the editor's own schema per story. It
stands on the document node under `doc.attrs.stories`, so an edit to it rides a transaction and
lands in the history, and everything the edit did not touch is written back as the bytes it arrived
as. Editing a comment used to flatten its body to plain text and rewrite it as one run, which lost
its bold, its paragraph style and its second paragraph.

`setCommentBody(id, body)` is a new command on `./commands`. It takes a `doc` node of `docxSchema`,
which is what a composer of your own builds a formatted body as; `updateComment(id, text)` stays
what it was for a body that is only text.

Dropping those attrs is what the story replaces them with: what a comment or a note says is no
longer written on the node marking where it stands, so nothing has to keep the two in step, a copy
of the node no longer carries the body along with it, and neither reaches the clipboard.
