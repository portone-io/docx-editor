---
"@portone/docx-editor": minor
---

Let a document be opened for comments alone, with each comment its author's own.

An editor now runs under a protection - everything, comments alone, or nothing - named after the OOXML `w:documentProtection` levels (`none`, `comments`, `readOnly`).
One guard answers for it, so an edit the protection shuts is refused however it was asked for, including a paste, a drag, an IME composition, and a consumer's own command run through `canRunCommand`.
The controls stay honest about it: a command that cannot apply reports so before it is run, which is what draws a button as unavailable rather than as one that swallows the click.
Read the level with the new `editingProtection` query.

A comment author now carries an identity beside the display name.
Hand one in with the comment and it is written to the `people` part of the file, under a provider id of this editor's own, and read back from it when the file is opened again.
An identity another application recorded, a signed-in Word user's above all, is one the editor cannot vouch for and is read as none.
Because a name is what both Word and this editor key that part by, a name that stands for two identities in one file resolves to none rather than handing one author's comments to the other.

Comments are then editable by whoever wrote them: a comment carrying an identity is its author's to edit, move and delete, and one carrying none stays open to everyone.
Replying and settling a thread as resolved or open belong to everyone, and a thread's root takes its replies with it.
An author's identity is never rewritten, not even by its owner.
Ask `canEditComment` what a control may offer, or open the editor for a moderator who may edit every comment there is.

Finally, a server taking a file back can hold the same rule over the bytes with `onlyCommentsChangedBy` from `@portone/docx-editor/core`, since the editor's own refusal is a courtesy the browser could go around.
It judges the whole package - every other part has to arrive as it left, the paper and the styles included - and answers `{ ok: true }` or `{ ok: false, reason }`, naming what it refused: a changed body, a comment that is not the author's, a forged author, a changed part, or a changed relationship.
Pass `{ editableComments: "all" }` for a file that came from a moderator's editor, where every comment was editable; an author's identity is refused under either setting.
