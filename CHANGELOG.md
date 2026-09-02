# @portone/docx-editor

## 0.2.0

### Minor Changes

- [#45](https://github.com/portone-io/docx-editor/pull/45) [`3190a35`](https://github.com/portone-io/docx-editor/commit/3190a359e0686bdc14166e22b1d96c761dabe251) Thanks [@Deea222](https://github.com/Deea222)! - Add a `comment` mode and record who wrote each comment.
  
  **Breaking:** `mode` is now required and the `commentAuthor` prop is removed.
  `author: { id, name, initials? }` moves into `mode` for its `comment` and `edit` kinds.
  `contextMenus` moves off `mode` onto `DocxEditor` itself, so `mode: { kind: "edit", contextMenus: false }` becomes `contextMenus={false}`.
  
  In `comment` mode the text can be selected and copied but not changed, while comments can be written, answered, resolved, and reopened.
  A comment records its author's identity in the document, so two authors sharing a display name stay distinct.
  Only a comment's author may edit or delete it, and anyone may reply, resolve, and reopen.
  `mode.editableComments: "all"` opens every comment to a moderator.
  The `canEditComment` and `editingProtection` queries report what the current mode allows.
  A server can check a file coming back with `onlyCommentsChangedBy` from `@portone/docx-editor/core`, which answers `{ ok: true }` or `{ ok: false, reason }` for whether the file differs in nothing but one author's comments.

## 0.1.1

### Patch Changes

- [#36](https://github.com/portone-io/docx-editor/pull/36) [`ab518b9`](https://github.com/portone-io/docx-editor/commit/ab518b9c551afc1d647ad612b265ec5895601240) Thanks [@Deea222](https://github.com/Deea222)! - Read two Unicode spellings of one font name as the same font.
  
  A name a document writes down can be composed or decomposed (NFC or NFD) while the same name elsewhere in the document, or among the toolbar's presets, is written the other way.
  Those spellings used to compare unequal, so the font dropdown listed the same font twice and a selection that all carried one font could report as mixed and blank the dropdown.
  Names are now compared in composed form, and the spelling the document wrote is what stays stored, exported, and shown.
  
  A list marker cut to its length cap is also cut between characters now, instead of possibly splitting an emoji or a combining sequence in half and drawing a replacement glyph.

## 0.1.0

The first release.

[The documentation](https://docx-editor.portone.io/docs) covers what the editor edits, what it preserves untouched, and its current limitations.
