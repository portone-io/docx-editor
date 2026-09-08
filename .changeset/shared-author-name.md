---
"@portone/docx-editor": patch
---

Take a comment added to a document whose author shares your display name, instead of refusing the whole save.
A document an older Word wrote has no people part, so its comments name an author and nobody in particular.
Adding a comment under one of those names no longer records a person for it, which had handed the comments already there to whoever commented next, and `onlyCommentsChangedBy` now takes such a comment back rather than reading it as a forged author.
Every comment carrying a shared name, the new one included, stays everyone's to edit.
This also protects comments preserved outside the editable story: their authors are not assigned to a new commenter, and the verifier rejects a people record that would claim them.
