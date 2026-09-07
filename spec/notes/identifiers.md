# Identifiers a document keeps unique

Several names in a WordprocessingML document may be held by one element only.
An edit that makes one node out of another - a split, or a copy dragged into place - can leave such a name standing in two places, and this note records what each name's specification asks for and how the export settles a conflict.

## Paragraph identifiers (`w14:paraId`, `w14:textId`)

`w14:paraId` is an `ST_LongHexNumber` (Part 1 §17.18.50: eight hexadecimal digits) that identifies a paragraph uniquely within its document part (Part 1 §11.3), except across the choices and fallback of one `mc:AlternateContent` block.
Its value shall be greater than 0 and less than `0x80000000`.
`w14:textId` is a version identifier for the paragraph's text in the same form and range, and an element carrying `w14:textId` shall carry `w14:paraId` as well.
Neither attribute is part of ECMA-376; both belong to Word's `w14` extension namespace.

Observed 2026-09-07 against Microsoft [MS-DOCX] revision 23.0 (2026-08-18), §2.6.2.3 `paraId` and §2.6.2.4 `textId`, read at learn.microsoft.com.

## Content control ids (`w:sdt` `w:id`)

A control's `w:id` is a unique numerical identifier that shall be persisted across sessions.
The specification settles a conflict itself: where several controls carry the same value, the first in the document keeps it and every later one is assigned a new identifier when the document is opened; a control carrying none is assigned one on opening.

Observed 2026-09-07 against ECMA-376 5th edition, Part 1, §17.5.2.18.
[Content controls](./contentControls.md#the-id-is-what-keeps-a-lock-from-fragmenting) records what the persistence clause means for a lock.

## Drawing object ids (`wp:docPr` `id`)

`wp:docPr` carries an `id` that identifies the DrawingML object uniquely within the document, so that other parts of the document can refer to it.

Observed 2026-09-07 against ECMA-376 5th edition, Part 1, §20.4.2.5.

## Bookmark ids (`w:bookmarkStart` `w:id`)

The `w:id` of a range marker is an annotation identifier, unique within the document, and it is what pairs a `w:bookmarkStart` with its `w:bookmarkEnd`: a start with no later end of the same id, or an end with no earlier start, makes the document non-conformant.
A pair copied whole therefore puts two starts under one id.

Observed 2026-09-07 against ECMA-376 5th edition, Part 1, §17.13.6.1 and §17.13.6.2.
[Bookmarks](./bookmarks.md) records how a marker is preserved.

## What we decide

The export settles every one of these names in one pass over the blocks, in document order, and the rule is the one §17.5.2.18 gives for controls: the first node to claim a name keeps it, and every later claimant yields.
Which node is first is its place in the document at export, not the order the edits happened in, so a copy dragged above its original is the one that keeps the name.

How a later claimant yields depends on the name:

- A block opened from the file goes back out as its original bytes only for the first node claiming that block; a later one is written from its own attributes.
- A later paragraph carrying a `w14:paraId` already written goes out without it, and without its `w14:textId`, which may not stand alone.
  The identifier is compared as the number it spells, so two spellings of one value are one name.
  A value outside what [MS-DOCX] allows is not an identifier to a reader either, and is left as it came rather than corrected.
- A later copy of a content control opens under a `w:id` of its own, which is what the specification would assign on opening.
- A block preserved as nothing but its original XML - a body-level bookmark marker or a body element the editor does not model - has nothing to be rewritten from, so a document holding one twice is refused with `unsupported-content` rather than written with two starts under one id or two copies of one section break.

A drawing copied inside the editor is not renamed yet; a newly inserted image takes an `id` above the highest the document holds.
