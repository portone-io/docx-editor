---
"@portone/docx-editor": minor
---

Add a `comment` mode, a reviewer's surface: the text may be selected and copied but not changed, while comments may be written, answered, resolved and reopened.
It is the standing OOXML document protection calls `comments`, beside the `readOnly` and `edit` modes the editor already had.
The mode may be changed on an open document; the editor, its history and the consumer's plugins stay.
`contextMenus` moves out of the mode and onto `DocxEditor` itself, beside `plugins`: the plugins behind it are read once when the editor mounts, so `mode: { kind: "edit", contextMenus: false }` becomes `contextMenus={false}` on the component and holds through every later mode.

The mode now names whose comments are written.
`mode: { kind: "comment" | "edit", author: { id, name, initials? } }` replaces the `commentAuthor` prop, and `id` is required: a reader (`{ kind: "readOnly" }`) names none.
Because of that, `mode` itself is required.

A comment records the identity it was written under in the document's `people.xml` part, the way Word records a signed-in author, so a host can tell two authors of one name apart and a file coming back can say who wrote what.
Editing and deleting a comment are its author's: a comment carrying another recognised identity is offered for replying and settling only.
A comment carrying no identity, such as one made in Word or before this release, stays open to everyone.
`mode.editableComments: "all"` opens every comment to a moderator.

For controls of your own, every command already answers `false` for what the mode refuses, and `editingProtection(state)` and `canEditComment(state, commentId, replyId?)` read the mode's standing.
A server taking the file back holds the same rule over the bytes with `onlyCommentsChangedBy(original, submitted, authorId)`, which answers `{ ok: true }` or `{ ok: false, reason }`.
