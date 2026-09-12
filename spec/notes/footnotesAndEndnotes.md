# Footnotes and endnotes

## Package and reference pairing

The Footnotes and Endnotes parts are related from the Main Document part and rooted at `w:footnotes` and `w:endnotes`. Internal relationship targets are resolved relative to the main part, including `.` and `..` path segments. A main-story `w:footnoteReference` or `w:endnoteReference` uses its required `w:id` to identify a `w:footnote` or `w:endnote` in the corresponding part. A missing target id makes the document non-conformant.

The parts also contain separator, continuation-separator, and continuation-notice entries. Those are layout instructions rather than document notes and are not shown in the note list.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§11.3.4, 11.3.7, 17.11.7, 17.11.14, and Part 2, §§6.4, 6.5.2.3.

## Reference order

Visible labels follow first-reference order within the main story rather than note-part order or raw ids, which commonly begin after reserved separator ids.
Footnotes and endnotes are counted separately, and a repeated id takes the label of its first reference.
The screen label follows `w:numFmt`, `w:numStart`, and a `w:numRestart` of `eachSect` from the section's own properties or else the settings; a restart on each page is counted straight through, since the label would then depend on the page layout it takes part in.
A format the editor has no speller for is drawn in decimal.
Where neither the section nor the settings names a format, endnotes are labelled in decimal just as footnotes are: Part 1 §17.11.17 and §17.11.18 give decimal for an omitted `w:numFmt`, and the editor keeps decimal for endnotes by product decision.
A reference carrying `customMarkFollows` draws no automatic inline label and takes no number, which §17.11.14 describes as not incrementing the count.
A reference naming a separator, continuation-separator, or continuation-notice entry takes no number either.

Observed 2026-09-12 against ECMA-376 5th edition, Part 1, §§17.11.4, 17.11.5, 17.11.11, 17.11.12, 17.11.14, 17.11.17-17.11.20, 17.18.59, 17.18.74.

## Editing boundary

Footnotes and endnotes are separate document stories rather than text owned by their main-story reference nodes.
Each body is read as such a story - a document of the editor's own schema, held on the document node under `footnote:<id>` or `endnote:<id>` - and `documentNotes` projects the text of it.
The Footnotes and Endnotes parts are written through the same story writer as a header part: an untouched entry goes back as its bytes, a removed story drops its entry, and a new story is appended after the part's own entries in id order.
Separator, continuation-separator, and continuation-notice entries go back as they arrived, and an export that changed one is refused.
A part the package lacked is created with a separator and a continuation-separator entry and no reference from `settings.xml`, which the schema leaves optional and which a Google Docs export also omits.
A note the editor inserts takes the id one above every id its own part holds or a reference of its kind names, separators and unreferenced entries included; the two parts are counted apart.
Deleting the last reference to a note deletes its story, and so its entry; an entry no reference named when the document opened is never deleted.
Copying a reference gives the copy a new id and a copy of the story without its paragraph ids, so the part holds no identifier twice.
A note edit reaches the document as a change to its story, so a lock at the reference and the editing protection judge it where they judge a body edit.
Word draws a note's number from the `w:footnoteRef` or `w:endnoteRef` its entry opens with, so an edit inside the note that removes that mark has it put back at the head of the first paragraph, and an emptied note still writes its number; `setFootnoteBody` and `setEndnoteBody` write the body as given and add no mark.
A footnote is drawn at the foot of the page its first reference stands on, and one that no longer fits is carried whole to the next page.
An endnote is drawn after the last block of the document, under one separator, onto further pages of the last section.
A `w:endnotePr/w:pos` of `sectEnd` is drawn as `docEnd`, the default.

Observed 2026-09-12 against ECMA-376 5th edition, Part 1, §17.11, and the Transitional schema's `CT_FtnEdn` and `CT_FtnDocProps`.

A note story is not comparable to a comment story for a protection: a comment protection lets the comment stories change and nothing else, so rewriting a note body is a body change in the editor and a changed part to the server verifier.
