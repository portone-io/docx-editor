---
"@portone/docx-editor": patch
---

The comments panel shows when each comment and reply was written in the reader's time zone and browser language, rather than in UTC.
`DocumentComment.date` and `DocumentCommentReply.date` are now that instant as an ISO 8601 string: the UTC time Word 2019 and later record for a comment where the file has one, and otherwise the comment's clock time read on the reader's clock, which is what Word shows.
A comment added with `NewComment.date` reads the same instant back, and the saved file records it the way Word does, so Word shows it in each reader's time zone as well.
A comment written by an earlier version of the editor recorded UTC time without that record, so it now reads as that UTC time, as it does in Word.
`onlyCommentsChangedBy` accepts the parts Word records a comment's time in, `commentsIds.xml` and `commentsExtensible.xml`, as the editor writes them, and refuses a returned file that removes an entry of a comment still in it from any comment part.
