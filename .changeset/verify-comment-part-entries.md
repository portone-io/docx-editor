---
"@portone/docx-editor": patch
---

Read the comment parts entry by entry in `onlyCommentsChangedBy`.

The three parts a comment is written across are the ones a comment edit may rewrite, so the package comparison passes over their bytes.
Nothing read them afterwards, which let a submission carry a field pointing at a remote image inside a comment body, a comment attributed to a third author that nothing refers to, or markup wrapped around a body, and still be answered as a file where only comments changed.
Each entry now has to arrive as it was, or be one this editor writes for an author who could have written it, and an entry nothing refers to has to stay as it was.
A file that fails is refused as `part-changed` naming the comment part.

Settling or replying to a comment that arrived with the file no longer rewrites its entry as plain text.
The entry keeps what it said, and its last paragraph gains the `w14:paraId` the thread state is written against.
The extended comments part carries an entry only for a comment that has thread state.
