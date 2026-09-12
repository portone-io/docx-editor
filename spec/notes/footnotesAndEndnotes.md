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
The Footnotes part is written through the same story writer as a header part: an untouched entry goes back as its bytes, a removed story drops its entry, and a new story is appended after the part's own entries in id order.
Separator, continuation-separator, and continuation-notice entries go back as they arrived, and an export that changed one is refused.
A Footnotes part the package lacked is created with a separator and a continuation-separator entry and no reference from `settings.xml`, which the schema leaves optional and which a Google Docs export also omits.
The Endnotes part is written the same way, under the names `w:endnotes`, `w:endnote`, and the Endnotes relationship and content type.
A note the editor inserts takes the id one above every entry its own part holds, separators and entries no reference names included, and above every id a reference of that kind names; the two parts are counted apart.
An edit that deletes the last reference to a note deletes its story, which is how an entry leaves the part; an entry no reference named when the document opened is never deleted this way.
An edit that copies a note reference gives the copy such an id and a copy of the story with its paragraph ids removed, so the part holds no identifier twice.
A note is edited where it is drawn, in one editor view mounted over that note alone, whose every change reaches the document as one story change, so a lock at the reference and the editing protection judge a note edit where they judge a body edit.
A note keeps the mark its entry opens with - the `w:footnoteRef` or `w:endnoteRef` its number is drawn from - through every edit inside the note: the mark arrives as a preserved fragment no deletion guard answers for, so an edit that carries it off has it put back at the head of the first paragraph that takes it, as the fragment it arrived as, and a note emptied of its text still writes its number.
An edit that writes a note body from outside that view, through `setFootnoteBody` or `setEndnoteBody`, writes the body as given and adds no such mark.
A footnote is drawn at the foot of the page its first reference stands on, in room the page layout keeps for it, and one that no longer fits is carried whole to the next page's room.
An endnote is drawn after the last block of the document, under a separator drawn once, onto pages of the last section as the endnotes run over.
A document whose `w:endnotePr/w:pos` asks for `sectEnd` is drawn at the end of the document all the same; only `docEnd`, the default, is honoured.

Observed 2026-09-12 against ECMA-376 5th edition, Part 1, §17.11, and the Transitional schema's `CT_FtnEdn` and `CT_FtnDocProps`.

A note story is not comparable to a comment story for a protection: a comment protection lets the comment stories change and nothing else, so rewriting a note body is a body change in the editor and a changed part to the server verifier.
