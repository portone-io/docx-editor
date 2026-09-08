---
"@portone/docx-editor": patch
---

The comment composer keeps the text it was opened over while the document is edited elsewhere, and closes when that text is deleted or the mode turns read-only. The comment is written on that text rather than on whatever is selected when the writer submits, and a refused write says why and keeps what was typed instead of leaving the form standing.

Opening the composer no longer scrolls the document back to its first page.

`addComment` and `canAddComment` take an optional `{ from, to }` range as their last argument, which a composer of your own can hold across edits the same way; the current selection is still used when none is given.
