---
"@portone/docx-editor": minor
---

Add a `comment` mode and record who wrote each comment.

**Breaking:** `mode` is now required and the `commentAuthor` prop is removed.
`author: { id, name, initials? }` moves into `mode` for its `comment` and `edit` kinds.
`contextMenus` moves off `mode` onto `DocxEditor` itself, so `mode: { kind: "edit", contextMenus: false }` becomes `contextMenus={false}`.

In `comment` mode the text can be selected and copied but not changed, while comments can be written, answered, resolved, and reopened.
A comment records its author's identity in the document, so two authors sharing a display name stay distinct.
Only a comment's author may edit or delete it, and anyone may reply, resolve, and reopen.
`mode.editableComments: "all"` opens every comment to a moderator.
The `canEditComment` and `editingProtection` queries report what the current mode allows.
A server can check a file coming back with `onlyCommentsChangedBy` from `@portone/docx-editor/core`, which answers `{ ok: true }` or `{ ok: false, reason }` for whether the file differs in nothing but one author's comments.
