---
"@portone/docx-editor": patch
---

Read the comment parts entry by entry in `onlyCommentsChangedBy`.

The three parts a comment is written across are the ones a comment edit may rewrite, so the package comparison passes over their bytes.
Nothing read them afterwards, which let a submission carry a field pointing at a remote image inside a comment body, a comment attributed to a third author that nothing refers to, or markup wrapped around a body, and still be answered as a file where only comments changed.
Each entry now has to arrive as it was, or be one this editor writes for an author who could have written it, and an entry nothing refers to has to stay as it was.
A file that fails is refused as `part-changed` naming the comment part.
