# @portone/docx-editor

## 0.2.1

### Patch Changes

- [#53](https://github.com/portone-io/docx-editor/pull/53) [`d264803`](https://github.com/portone-io/docx-editor/commit/d264803381c3272ae31e1af1faa7fd158b95a81f) Thanks [@Deea222](https://github.com/Deea222)! - Stop text copied out of the editor from carrying the document's private data.
  
  A copy was drawn the way the editor draws itself, so the HTML it left on the clipboard held the paragraph and run XML, the name and recorded identity of a comment's author, what the comment and its replies say, and the body of a footnote.
  All of it landed in whatever application it was pasted into, and in `comment` mode a reader who may not change the body could take the body out this way.
  A copy now carries only what a reader of it needs, and the paragraph style it names is the style's id rather than the whole `w:pPr`.
  A link goes out as a link, so it can be followed where it lands and comes back as one when it is pasted here again.
  
  The plain text beside it says what was copied: a tab is a tab, a line break is a line, a page break is a form feed, and a table's cells stand apart by a tab and its rows by a line, so a table pasted into a spreadsheet arrives as a table.
  A cell holding a line break is the exception, since the line inside it reads as the start of the next row.

- [#50](https://github.com/portone-io/docx-editor/pull/50) [`41d82c5`](https://github.com/portone-io/docx-editor/commit/41d82c5bee74597baedd813c43c8dcfcdfe74fbf) Thanks [@Deea222](https://github.com/Deea222)! - Stop `onlyCommentsChangedBy` from excusing a part a submission relates as a comment part.
  
  The three comment parts are left out of the byte comparison, and which parts those were came from the submitted file's own relationships.
  A file could relate a second comments, extended comments or people part at any part it liked and have that part go uncompared, so a rewritten styles part, a settings part pointing at a template off the package, replaced image bytes or forged document properties were all reported as a change to nothing but comments.
  The parts left out are now the ones the reader opened, a comment part may be related once, one related for the first time has to be a part the submission brought with it, and a relationship part naming one id twice is turned down.
  
  A file that relates an extended comments part but carries no comments part now has that part read, so writing the first comment into it writes that part rather than a second one beside it.
  Two extended parts related at once was a file whose settled threads a reader would lose, since only the first of them is read.
  
  This affects 0.2.0. A server that accepted files on this verdict should upgrade and run the check again over what it accepted, where such a file now answers `part-changed` naming the part, or `relationship-changed` naming the relationship part.

- [#56](https://github.com/portone-io/docx-editor/pull/56) [`467f6f5`](https://github.com/portone-io/docx-editor/commit/467f6f5781dbf6eb6739dbf3369155c7287ee608) Thanks [@Deea222](https://github.com/Deea222)! - Commands and `canRunCommand` now report false where a bookmark marker or a note reference would be removed, instead of reporting true and changing nothing.
  
  A refusal over one of those markers also ends an open IME composition, the way a refusal over a locked control already did.

- [#59](https://github.com/portone-io/docx-editor/pull/59) [`6d7b97f`](https://github.com/portone-io/docx-editor/commit/6d7b97fdb89e0aba7ac13057b503ec510a5550c1) Thanks [@Deea222](https://github.com/Deea222)! - Formatting XML is now written through shared primitives that consistently escape attribute values while preserving existing formatting behavior.

- [#58](https://github.com/portone-io/docx-editor/pull/58) [`d9edb75`](https://github.com/portone-io/docx-editor/commit/d9edb75fa45e1a6cc81e5741a5242cf500107091) Thanks [@Deea222](https://github.com/Deea222)! - Pagination measures every block as a list of break candidates; no visible change

- [#55](https://github.com/portone-io/docx-editor/pull/55) [`7110baa`](https://github.com/portone-io/docx-editor/commit/7110baae68b75323720a39374c6df58db4367896) Thanks [@Deea222](https://github.com/Deea222)! - Turn down a raw OOXML fragment that does not hold its shape as it enters the document, rather than writing it into the exported file.
  
  Keep imported formatting when a document uses an inherited default namespace or an alternative WordprocessingML prefix. Reject nested namespace rebinding that could disconnect hyperlinks from their targets on export.
  
  The editor draws the original XML of a paragraph, a run, a table, an image or an annotation into the page as a `data-` attribute, so that it can read the live DOM back after an IME composition or a browser edit. A fragment written into one of those attributes from outside, by a consumer plugin or through `view.pasteHTML`, is now held to what that attribute goes back out as: a whole element under an expected name, an opening tag's attributes, or an opening tag the writer closes itself. One that does not hold it is turned down along with the rule reading it, so the content settles one level plainer - a paragraph keeps its text and loses its properties, a run loses its mark, a content control or a hyperlink loses its wrapper - instead of a `data-ppr` reading `</w:p><w:p>...` writing a second paragraph into the exported body, or one that never closed reaching export and failing the whole document with `malformed-xml`.

- [#54](https://github.com/portone-io/docx-editor/pull/54) [`8aea477`](https://github.com/portone-io/docx-editor/commit/8aea477dceab5462c15903b25b29c738980aff47) Thanks [@Deea222](https://github.com/Deea222)! - An untouched table opened in the editor is written back byte for byte, its `w:tblGridChange` included.
  
  Export decided whether a block was untouched by comparing it with the one import produced, attrs and all.
  Opening a document works some of those attrs out again from the formatting around them - a table's shared cell borders among them - so a table nobody had touched compared unequal and was rebuilt, and the rebuild dropped the markup the writer does not model.
  
  Each attr now declares whether the writer writes from it, whether the editor works it out for the screen, or whether it identifies something in the open document, and the comparison ignores the ones worked out for the screen.
  Exporting through `createEditorState` is held to the same byte identity as exporting straight from `importDocx`.

- [#48](https://github.com/portone-io/docx-editor/pull/48) [`de84006`](https://github.com/portone-io/docx-editor/commit/de840061dd60f3b652273410bc2d0a85b9451ad3) Thanks [@Deea222](https://github.com/Deea222)! - Stop `onlyCommentsChangedBy` from refusing a comment written into a table cell.
  
  A commented table is rebuilt on the way out, so it came back worded the way this editor words it while the original carried the wording its producer chose, and the two were compared word for word.
  The story is now compared as this editor writes it back, so two blocks this editor would write alike are read alike: the attribute order inside a `w:tcW` or a `w:tblW`, a percentage width written as `100%` or as `5000`, runs a producer split that say the same text, and a table's `w:tblGridChange`.
  Formatting properties keep the line breaks, comments and text a producer wrote between them, so rewriting a run's bold, a paragraph's alignment or a cell of a rebuilt table no longer drops them.
  A rebuilt table keeps what stood inside its width and span properties too.

- [#52](https://github.com/portone-io/docx-editor/pull/52) [`991ba48`](https://github.com/portone-io/docx-editor/commit/991ba48501e20e60bfc69b3c76862f4a26123b3e) Thanks [@Deea222](https://github.com/Deea222)! - Read the comment parts entry by entry in `onlyCommentsChangedBy`.
  
  The three parts a comment is written across are the ones a comment edit may rewrite, so the package comparison passes over their bytes.
  Nothing read them afterwards, which let a submission carry a field pointing at a remote image inside a comment body, a comment attributed to a third author that nothing refers to, or markup wrapped around a body, and still be answered as a file where only comments changed.
  Each entry now has to arrive as it was, or be one this editor writes for an author who could have written it, and an entry nothing refers to has to stay as it was.
  A file that fails is refused as `part-changed` naming the comment part.
  
  Settling or replying to a comment that arrived with the file no longer rewrites its entry as plain text.
  The entry keeps what it said, and its last paragraph gains the `w14:paraId` the thread state is written against.
  The extended comments part carries an entry only for a comment that has thread state.

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
