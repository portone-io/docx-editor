---
"@portone/docx-editor": patch
---

The comments panel shows when each comment and reply was written in the reader's time zone and browser language, rather than in UTC.
`DocumentComment.date` and `DocumentCommentReply.date` are now that instant as an ISO 8601 string: the UTC time Word 2019 and later record for a comment where the file has one, and otherwise the comment's clock time read on the reader's clock, which is what Word shows.
A comment added with `NewComment.date` reads the same instant back, and the saved file records it the way Word does, so Word shows it in each reader's time zone as well.
A comment written by an earlier version of the editor recorded only a UTC time, so its UTC clock digits are now shown as the reader's local time, as Word shows them: it appears shifted by the reader's UTC offset.
`onlyCommentsChangedBy` accepts the parts Word records a comment's time in, `commentsIds.xml` and `commentsExtensible.xml`, as the editor writes them, and refuses a returned file that removes an entry of a comment still in it from any comment part.
Put the server running `onlyCommentsChangedBy` on this release or later before, or together with, the editor, as [the verifier's version rule](https://docx-editor.portone.io/docs/core/verifying-a-commenters-file#how-files-are-compared) asks: an earlier verifier refuses every file in which this editor wrote a new comment.
