# @portone/docx-editor

## 0.4.0

### Minor Changes

- [#102](https://github.com/portone-io/docx-editor/pull/102) [`4258530`](https://github.com/portone-io/docx-editor/commit/4258530aaf391ba546a282f73dff7b45f7063829) Thanks [@Deea222](https://github.com/Deea222)! - Draw the table styles and list numbering a document defines, and let a document that never held a list take one.
  
  A table style is now displayed part by part: the header row, the closing row, the first and last column, the corners and the banded rows take the shading, the lines and the text formatting the style dresses them with (`w:tblStylePr`), and `w:tblLook` decides which of those parts a table takes.
  The bands are as many rows or columns wide as the style says (`w:tblStyleRowBandSize`, `w:tblStyleColBandSize`), and the header row and the closing row are not banded with the rest.
  Text already in a cell follows its new position after a row edit, and the conditional-formatting markup a document arrived with goes back out untouched.
  
  A list that takes its numbers from a numbering style now draws them.
  Such a list holds no numbers of its own: it names a style (`w:numStyleLink`), the style names a list, and that list's definition is where the numbers are, which is how Word writes a list built from its gallery.
  Alongside decimal, bullets, upper and lower letters and lower Roman numerals, a list may now count in upper Roman numerals, decimal numbers with a leading zero, Ganada, Korean digits and the Chinese counting system; a format past those is still shown with decimal numbers.
  Each level is drawn the way it asks to be: where its counting starts over (`w:lvlRestart`), whether the numbers in its text are all spelled as decimals (`w:isLgl`), what stands between its number and the paragraph text (`w:suff`), and where the number sits in the room kept for it (`w:lvlJc`).
  A marker is also drawn in the character formatting its own level writes down (`lvl/rPr`), so bold, italic, color, size and typeface reach the number and never the text of the paragraph it stands in front of.
  All of this is read; the numbering part goes back out exactly as it arrived.
  
  A new list no longer needs the document to have arrived with a numbering part.
  The export writes `word/numbering.xml`, relates it from the main part and declares it in the content types, so a document that never held a list can take one; the list commands, which used to report that they did not apply in such a document, now do.
  If an existing numbering relationship points to a missing part, the new definition is written at that location so the document can find the list when it is reopened.
  
  A new or pasted list is exported with the definition it was started with, rather than a format inferred from its numbering IDs, and undo restores that registration along with the edit.
  A list whose definition is missing or unsupported is refused as `unsupported-content`, and a package with no `[Content_Types].xml` for the new numbering part to be declared in is refused as `missing-content-types`; `exportProblems` reports either ahead of the write.
  List and abstract definition IDs stay distinct in a document already using the largest safely representable integer, and a value that cannot be written faithfully is rejected rather than silently changed or discarded.
  
  An ECMA-376 Strict package is refused with the new `unsupported-conformance` code, and a document whose main part root does not bind `w` to Transitional WordprocessingML is refused as `unsupported-content`.
  Word saves neither by default, and both used to open far enough to fail later and obscurely: a Strict package as `missing-part`, and a document under another prefix as a `malformed-xml` about bookmarks on its first edit.
  A write needing a prefix already bound to a different namespace is refused instead of silently giving new markup the wrong meaning, and a new hyperlink relationship attribute is checked for a declaration that shadows the root binding.
  The comment and hyperlink writers now have the part's root declare the prefix they write under rather than declaring it on every element written, and editing an existing numbering or comment extension part that uses another prefix adds the writer's declaration at the root, so an untouched document still exports byte for byte and the exported XML stays readable.
  
  `DocxImportErrorCode` gains `unsupported-conformance`.
  `DocxExportErrorCode` loses `missing-numbering-part`, since nothing can reach it any more, and that is a compile error for a `switch` over the code written to be exhaustive; drop the branch.
  `missing-content-types` now also covers the numbering part a new list is defined in.
  On `./core`, `Numbering` gains `added`, and the new `NewList` and `NewListLevel` types describe the definitions the editor registered.
  Level maps are read-only, and a registered definition carries a restart, legal numbering and a suffix, while marker run formatting and custom tab stops stay outside what can be registered.

### Patch Changes

- [#101](https://github.com/portone-io/docx-editor/pull/101) [`1b6c723`](https://github.com/portone-io/docx-editor/commit/1b6c723982eed9c10656d57aca1d496e6baad5ec) Thanks [@Deea222](https://github.com/Deea222)! - Stop drawing the page number on the corner of each page.
  
  The page guides laid a small grey number inside the top right corner of every page, over the paper the text sits on.
  It is gone, and the guides now draw the gaps between pages and the header and footer stories alone.
  A document that prints its own page number through a `PAGE` field in a header or footer is unaffected: that number is the document's, not the editor's.
  
  The `docx-editor-page-badge` class the number carried is no longer emitted, so a rule of your own written against it in the published stylesheet no longer matches anything.

## 0.3.0

### Minor Changes

- [#62](https://github.com/portone-io/docx-editor/pull/62) [`5c481e7`](https://github.com/portone-io/docx-editor/commit/5c481e7c6e9a27373dad90d866a70cd80919b294) Thanks [@Deea222](https://github.com/Deea222)! - `onlyCommentsChangedBy` answers `comment-markup-rejected` where it used to answer `part-changed` for a comment part entry this editor would not have written for this author, or an entry nothing refers to that changed. `part` still names the comment part the entry sits in.
  
  The verdicts themselves are unchanged: every file accepted before is accepted now, and every file refused before is refused now. What moves is the name a server logs, so "a part this file was not supposed to touch" and "markup forged into a comment" no longer read alike. A `switch` over `verdict.reason` needs the new case.

- [#77](https://github.com/portone-io/docx-editor/pull/77) [`a0b5cc6`](https://github.com/portone-io/docx-editor/commit/a0b5cc643bbeb6cf3407a0d729e7cc74fec304d6) Thanks [@Deea222](https://github.com/Deea222)! - Ask whether a document can be exported before trying.
  
  `exportProblems(doc, session)` on the core entry, and `canExport(state)` with `documentExportProblems(state)` on the commands entry, report known reasons the writer would refuse the document, in the order it would raise them: each under the code and with the message the `DocxExportError` would carry, and with its position where the problem stands in the document. `exportDocx` throws the first entry of the same list, so problems reported by the query are also refused by the write.
  
  The editor's handle answers the same list as `exportProblems()` beside `exportBytes()`, and `downloadDocx` asks it first and returns `{ status: "blocked", problems }` instead of throwing. A refusal the list does not foresee, such as a node a plugin stripped of an attribute the writer needs, is still thrown. `DownloadDocxResult` gaining a fourth member is a compile error for a `switch` over `status` written to be exhaustive; add a `blocked` branch that shows the problems.
  
  A document whose comments part arrived as an empty element used to refuse its first comment with `malformed-xml`, since the writer looked for a closing tag the element does not have; the part is now opened for the entry, the way an empty extended comments part already was.

- [#63](https://github.com/portone-io/docx-editor/pull/63) [`1a4ec92`](https://github.com/portone-io/docx-editor/commit/1a4ec92162477d468fdb14a95a47388befba45f9) Thanks [@Deea222](https://github.com/Deea222)! - Read what a document holds that this editor cannot model. `importDocx` returns a `notes` array alongside the document and the session, the new `exportDocxReport` returns the same list beside the bytes it writes, and the new `documentFidelity(state)` on `./commands` answers the question about the document standing in an editor. A `FidelityNote` says how much of the original survived, what kind of content it was, the part and body block it came from, where in the document it stands, and the original element name, so a host can list what a file lost instead of guessing at it. Destructuring `importDocx` is unaffected.
  
  A table now carries its `w:tblGridChange` instead of losing it whenever the table is rebuilt. The grid is still written from the column widths, and the revision markup closes it where CT_TblGrid takes it. `onlyCommentsChangedBy` therefore catches a submission that lost a `w:tblGridChange`, where it used to accept one: a returned file that dropped it now answers `body-changed` rather than `ok`.

- [#79](https://github.com/portone-io/docx-editor/pull/79) [`70d4801`](https://github.com/portone-io/docx-editor/commit/70d48015f52a126824f25bbd03a65b5509699de2) Thanks [@Deea222](https://github.com/Deea222)! - Applying a paragraph style keeps the spacing the document defaults lay down, and formatting a style or the document defaults switch on can be switched off: the off is written into the run (`w:b w:val="0"`) and drawn as off, where it used to be dropped and the style's value drawn again.
  
  Every display value a paragraph or a run carries is now resolved in the ECMA-376 §17.7.2 order by one resolver, whichever path built the paragraph: opening the document, applying a style, a paragraph edit, a paste, or typing into a fresh paragraph. A character style a run points at (`w:rStyle`) takes its place in that order, a hanging indent's implicit tab stop follows the indent actually drawn, and a document's `w:noTabHangInd` setting switches that stop off.
  
  `RunFormat` grows to say so: `bold`, `italic`, `strike` and `smallCaps` are `boolean` (false is a toggle the run switches off outright), and `underline` may be `"none"`. A reader checking `=== true` or a truthy value is unaffected.
  
  Supported character defaults now appear in the text and toolbar, so a default bold setting turns off with one press. Clearing a direct font size immediately restores the inherited paragraph or character style in the formatting controls.

- [#82](https://github.com/portone-io/docx-editor/pull/82) [`f804ba1`](https://github.com/portone-io/docx-editor/commit/f804ba1eb5e5e6ed5ebf811e1edf2fa1858dc601) Thanks [@Deea222](https://github.com/Deea222)! - A paragraph marked keep with next (`w:keepNext`), by its own properties or by its style, stays on the same page as the start of the block after it in the page guides, as it does in Word. A run of such paragraphs moves together with the first piece of the block the keeps end at; a run no page can hold is laid out as if no keep were set. The document is not changed: the mark is read, never written.
  
  `ParagraphFormat` grows `keepNext?: boolean` to say so, beside `pageBreakBefore`.
  
  An explicit off (`w:keepNext w:val="0"`) overrides an inherited keep, so a paragraph can opt out of the keep imposed by its style.

- [#80](https://github.com/portone-io/docx-editor/pull/80) [`a37b4fe`](https://github.com/portone-io/docx-editor/commit/a37b4fea75bbf6e9c342c830d358ba21119dd17c) Thanks [@Deea222](https://github.com/Deea222)! - Run formatting is read, written and compared through one property table, so the value a control reads, the XML an edit writes and the check that leaves text already in that state alone can no longer drift apart. The XML written is byte for byte what it was.
  
  `RunFormat` gains `caps`, `doubleStrike` and `characterSpacingPt`, read off a run's `w:caps`, `w:dstrike` and `w:spacing`. They reach the `format` attr a plugin reads; the editor does not draw or edit them yet.
  
  An underline setter now distinguishes underline kinds; the public underline toggle still turns any existing kind off.

- [#72](https://github.com/portone-io/docx-editor/pull/72) [`578e73e`](https://github.com/portone-io/docx-editor/commit/578e73ed9444f528231048621a13651511438e79) Thanks [@Deea222](https://github.com/Deea222)! - `srcId` now names the block within the session it was opened in. A preserved block used to carry a bare index into the blocks of whichever document it was exported against, so a block moved or pasted in from another document pointed at this document's block of the same number and went out as that block's XML instead of its own. The attribute is now a string naming the document, the story and the place in it, and a block whose document is not the one being exported is refused with `lost-original` rather than written as something else.
  
  A plugin that read `node.attrs.srcId` as a number needs to change: it is a string, and the only thing to do with it is hand it back as it was found. Nothing else about the attribute is public, and no exported type or function signature changed.

- [#60](https://github.com/portone-io/docx-editor/pull/60) [`1b5d0fd`](https://github.com/portone-io/docx-editor/commit/1b5d0fd96f13341de0fe34374c2a7dedec57525c) Thanks [@Deea222](https://github.com/Deea222)! - Hand `importDocx`, `exportDocx`, `onlyCommentsChangedBy`, `documentNumbering` and `parseNumbering` an `xmlParser` to read a document on a runtime that has no `DOMParser` global, instead of installing one.
  
  A call given neither is refused with `DocxImportError` and the new import code `no-xml-parser`. It used to fail with a bare `ReferenceError`, which a server checking a file a counterparty returned could not tell apart from a document that arrived damaged. Reading a document no longer asks for a `Node` global at all, so `DOMParser`, however it is supplied, is the only thing the core entry needs from a DOM.
  
  Each entry point settles its parser as the call comes in, so a runtime holding none is turned down before the bytes are looked at: bytes that are not a docx opened without a parser now report `no-xml-parser` where they reported `not-a-docx`. A parser that answers markup it cannot read by throwing, rather than by handing back a document holding a `parsererror`, is read as `malformed-xml` instead of having its own exception reach the caller.

### Patch Changes

- [#75](https://github.com/portone-io/docx-editor/pull/75) [`c79e417`](https://github.com/portone-io/docx-editor/commit/c79e417bc7c2ec9cc3be804a1010533780558954) Thanks [@Deea222](https://github.com/Deea222)! - Document which schema attributes plugins may rely on and which raw OOXML attributes are internal.
  An internal classification now records the provenance of every node and mark attribute without changing import, editing, or export behavior.

- [#65](https://github.com/portone-io/docx-editor/pull/65) [`a9a7e6d`](https://github.com/portone-io/docx-editor/commit/a9a7e6d92277f7725d995415da43ae5f3488e079) Thanks [@Deea222](https://github.com/Deea222)! - Content control properties are written in the order the schema lays down. Locking a control that declares what kind of control it is - a date picker, a drop-down, plain text - used to write `w:lock` after that declaration, where CT_SdtPr puts `w:lock` before it, so a validator reading the exported file could refuse a control this editor had rewritten. A control carrying `w:label` or `w:tabIndex` was written the same wrong way round.
  
  Untouched documents retain their original XML.

- [#73](https://github.com/portone-io/docx-editor/pull/73) [`ddda545`](https://github.com/portone-io/docx-editor/commit/ddda545d43788007cf8c2c0507a33b1326d593ca) Thanks [@Deea222](https://github.com/Deea222)! - Enter no longer duplicates paragraph identifiers or a paragraph-level section break. The continuing paragraph keeps its identifiers, and the section break stays on the last paragraph of the split.
  
  Deleting or joining away a section-ending paragraph is refused until section editing is supported. Replacing text within that paragraph remains available.

- [#81](https://github.com/portone-io/docx-editor/pull/81) [`94d4ed7`](https://github.com/portone-io/docx-editor/commit/94d4ed748b22f49054acfb0b9cf7abcd83f19f09) Thanks [@Deea222](https://github.com/Deea222)! - The lines of a table's cells and the style values of a paragraph are worked out again by one plugin, which also works every value out again when the document's formatting is replaced under it; no behavior change for a document opened today.
  
  Changing the formatting context also refreshes existing run marks, so text does not keep its previous style after the paragraph display values change. These updates preserve the original run XML, including in locked content and protected documents.

- [#74](https://github.com/portone-io/docx-editor/pull/74) [`c22f1da`](https://github.com/portone-io/docx-editor/commit/c22f1da7e1aec0a4a5ffbd805eb7b761655ea84f) Thanks [@Deea222](https://github.com/Deea222)! - Commands now share guard helpers so their applicability checks and dispatched edits respect the same rules. Formatting at a caret inside locked content is refused; formatting queries continue to report the selected text's values under document protection.

- [#90](https://github.com/portone-io/docx-editor/pull/90) [`3b37fbd`](https://github.com/portone-io/docx-editor/commit/3b37fbd69eb90fd27011af4ecce2dfd47b61b64f) Thanks [@Deea222](https://github.com/Deea222)! - Importing `emuToPx`, `pxToEmu` or `toImageExtent` off `./core`, or the image file helpers off `./commands`, no longer carries the XML naming layer into a consumer's bundle. The picture module read the namespace table at the top of the file, which a bundler keeps as a side effect, so one multiplication cost 817 bytes minified where it costs 182. Nothing written into a document changes.

- [#70](https://github.com/portone-io/docx-editor/pull/70) [`9d7f659`](https://github.com/portone-io/docx-editor/commit/9d7f6598c2f010ee53753cbd3c7952e8b4bc7f1e) Thanks [@Deea222](https://github.com/Deea222)! - The editor reads document-level values such as styles, numbering and page geometry from one snapshot; no visible change.

- [#69](https://github.com/portone-io/docx-editor/pull/69) [`4f8c94a`](https://github.com/portone-io/docx-editor/commit/4f8c94abf38c4b77eecb632dc783f215d5decb41) Thanks [@Deea222](https://github.com/Deea222)! - A page measurement applies its pushes, break spaces and table continuations in one transaction instead of three. The pages look exactly as they did; what changes is that a single measurement now reaches the editor as a single state change, so anything watching transactions - an `onStateChange` handler, a plugin, a React state hook - sees one rather than three per remeasure.
  
  A table's repeated header now also refreshes as soon as its source row is edited, without waiting for the next measurement.
  
  Changing a continued row's formatting keeps its page gap until remeasurement. Page pushes also update when their measured contribution changes but the total top margin stays the same.

- [#83](https://github.com/portone-io/docx-editor/pull/83) [`14f3b87`](https://github.com/portone-io/docx-editor/commit/14f3b875004ae28c41a3fd0d58c7783fa9731d67) Thanks [@Deea222](https://github.com/Deea222)! - A relationships part whose root carries a namespace prefix, or that arrived as an empty element, is now spliced correctly when the export adds a relationship to it: the entry goes inside the root, spelled under the root's own prefix, where the export used to refuse the file with `malformed-xml` for want of a bare closing tag. A numbering part that arrived as an empty element takes a new list definition the same way instead of being refused.
  
  When one export adds several parts, `[Content_Types].xml` now declares them in the order they were added, a media type's `Default` ahead of any `Override`; each writer used to put its own declaration first, so the declarations came out in the reverse order of adding. Every declaration is still placed right after the opening tag, and the rest of the part is left as it arrived.
  
  Every XML part the export rewrites is now read back before the file is repacked, so a part that would not open is refused with `malformed-xml` naming the part rather than handed back.
  
  New part names and content-type requests recognize names differing only in case, avoiding duplicate package entries and declarations.

- [#71](https://github.com/portone-io/docx-editor/pull/71) [`a0bcd87`](https://github.com/portone-io/docx-editor/commit/a0bcd8798a33ecf3aba267bc2cdfbbca2b98d632) Thanks [@Deea222](https://github.com/Deea222)! - Comment-only verification and export now share the definitions for comment package parts. Public types are unchanged.
  
  Verification rejects new or altered content around comment entries, including outside the XML root, while accepting annotations preserved from the original and those removed by normal comment-part rewrites. Namespace rebindings under rewritten comment markup are also rejected.

- [#76](https://github.com/portone-io/docx-editor/pull/76) [`46c8ba7`](https://github.com/portone-io/docx-editor/commit/46c8ba73cedd06cd85141ea44f09119648315396) Thanks [@Deea222](https://github.com/Deea222)! - Table pagination moves into a block-kind module; no visible change. What the page engine knows about a table - where it may be parted between rows, the spacer and repeated header a continued page is drawn with, and which positions a page cut may stand at - used to be spread over the measurer, the decorations and the plugin state. It is now one module beside the paragraph's, and the engine asks whichever kind claims a block. Pages, page breaks and continued tables look exactly as they did.

- [#63](https://github.com/portone-io/docx-editor/pull/63) [`1a4ec92`](https://github.com/portone-io/docx-editor/commit/1a4ec92162477d468fdb14a95a47388befba45f9) Thanks [@Deea222](https://github.com/Deea222)! - Keep namespace bindings declared on a table grid when preserving its revision history, so editing the table can still produce a readable DOCX. Rebuilt tables also retain the required empty table-properties element when no properties are set.

- [#78](https://github.com/portone-io/docx-editor/pull/78) [`3bdf81d`](https://github.com/portone-io/docx-editor/commit/3bdf81dc3c6b0a92091e46fc8312907e60cfc67e) Thanks [@Deea222](https://github.com/Deea222)! - A paragraph copied inside the editor no longer exports a duplicate `w14:paraId`. The copy goes out as a paragraph of its own, without the original's identifiers, and the original still goes out as the bytes it arrived as.
  
  A preserved block copied twice, a body-level bookmark marker or a section break among them, is refused with `unsupported-content` instead of written twice. `exportProblems` and `canExport` report that refusal ahead of the write, at the place the second copy stands.

- [#67](https://github.com/portone-io/docx-editor/pull/67) [`77a468f`](https://github.com/portone-io/docx-editor/commit/77a468f5c4159e619a20b446bd35d14b4833f44a) Thanks [@Deea222](https://github.com/Deea222)! - Universal measures such as `8.5in` are read correctly. A measurement in a document may be written as a length with a unit - `8.5in`, `2.54cm`, `12pt` - as well as a count, and every one of them used to be read as its leading digits alone. A US Letter document whose section says `w:pgSz w:w="8.5in"` was drawn as a page 8.5 twips wide, which is no page at all, so the editor fell back to A4 and showed the wrong paper; a tab stop at `1.5in` landed at 0.08pt and an automatic tab interval of `0.75in` collapsed every tab in that document to no width. Page size and margins, indents, spacing, font sizes, table and cell widths, cell margins, row heights and tab stops are now read as the lengths they name.
  
  `on` and `off` spellings of boolean attributes are read as the schema admits. A style marked `w:default="on"` is now recognised as the default style for its kind, as `w:default="1"` already was, so a document that marks its defaults that way is shown with the formatting they lay down.
  
  Integer measurements retain support for an explicit `+` sign. Table and cell widths with an explicit `%` follow Word's percentage interpretation even when their width type says otherwise. Rebuilt table widths and grid columns are written in whole units; untouched XML is preserved.
  
  Overflowing universal-measure and percentage conversions are rejected during import.

- [#65](https://github.com/portone-io/docx-editor/pull/65) [`a9a7e6d`](https://github.com/portone-io/docx-editor/commit/a9a7e6d92277f7725d995415da43ae5f3488e079) Thanks [@Deea222](https://github.com/Deea222)! - An edit to a cell border, a cell shading, a row height, a paragraph indent, or a line spacing now reads and writes the WordprocessingML attribute alone. A producer's own attribute that shares the local name (`x:val` beside `w:val`, declared ignorable) used to be taken for the formatting value, or written over in its place; it is now left as the producer wrote it, and the value Word reads is the one read and written.

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
