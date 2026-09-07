---
"@portone/docx-editor": minor
---

`onlyCommentsChangedBy` answers `comment-markup-rejected` where it used to answer `part-changed` for a comment part entry this editor would not have written for this author, or an entry nothing refers to that changed. `part` still names the comment part the entry sits in.

The verdicts themselves are unchanged: every file accepted before is accepted now, and every file refused before is refused now. What moves is the name a server logs, so "a part this file was not supposed to touch" and "markup forged into a comment" no longer read alike. A `switch` over `verdict.reason` needs the new case.
