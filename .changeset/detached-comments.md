---
"@portone/docx-editor": minor
---

A comment now outlives the text it was written for, and a resolved one stops marking the page.

Deleting commented text used to delete the comment and everything said under it. The thread now
stays. Trim the highlighted text - type over part of it, backspace through its edge - and the
comment follows what is left. Delete all of it and the comment comes back detached: no longer drawn
beside the page, since there is nothing to draw it next to, and listed under All comments marked
"Original content deleted". Cutting commented text leaves it behind the same way instead of dropping
it. Word has no detached state, so such a comment reaches the saved file as an ordinary comment
sitting where the deleted text was. Deleting a comment from its own Delete button still deletes it,
and undo and redo still take a deletion back and make it again.

Resolving a thread now takes its highlight off the text, and reopening puts it back. Clicking a
resolved comment's card in All comments selects the text it was written about and scrolls it into
view, which `selectComment` now does for every caller.

`documentComments` reports a new `anchored` field on every comment, `false` for one that no longer
marks a stretch of the document. Its existing fields are unchanged, so a panel of your own keeps
working and can read the new field to decide what a click on a card should do.
