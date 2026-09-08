# Footnotes and endnotes

## Package and reference pairing

The Footnotes and Endnotes parts are related from the Main Document part and rooted at `w:footnotes` and `w:endnotes`. Internal relationship targets are resolved relative to the main part, including `.` and `..` path segments. A main-story `w:footnoteReference` or `w:endnoteReference` uses its required `w:id` to identify a `w:footnote` or `w:endnote` in the corresponding part. A missing target id makes the document non-conformant.

The parts also contain separator, continuation-separator, and continuation-notice entries. Those are layout instructions rather than document notes and are not shown in the note list.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§11.3.4, 11.3.7, 17.11.7, 17.11.14, and Part 2, §§6.4, 6.5.2.3.

## Reference order

Visible labels follow first-reference order within the main story rather than note-part order or raw ids, which commonly begin after reserved separator ids. Exact numbering formats, custom marks, and section restart rules are preserved in OOXML but are not reproduced by the screen label. A reference carrying `customMarkFollows` does not draw an automatic inline label.

## Editing boundary

Footnotes and endnotes are separate document stories rather than text owned by their main-story reference nodes. Each body is now read as such a story - a document of the editor's own schema, held on the document node under `footnote:<id>` or `endnote:<id>` - and `documentNotes` projects the text of it. What is still missing is the other half: the note parts are repacked unchanged, so nothing writes an edited note body back, and no command or surface offers editing one. Adding that means writing the note parts through the same story writer the Comments part goes through, and integrating a note's placement with document layout.

A note story is not comparable to a comment story for a protection: a comment protection lets the comment stories change and nothing else, so rewriting a note body is a body change in the editor and a changed part to the server verifier.
