# Footnotes and endnotes

## Package and reference pairing

The Footnotes and Endnotes parts are related from the Main Document part and rooted at `w:footnotes` and `w:endnotes`. Internal relationship targets are resolved relative to the main part, including `.` and `..` path segments. A main-story `w:footnoteReference` or `w:endnoteReference` uses its required `w:id` to identify a `w:footnote` or `w:endnote` in the corresponding part. A missing target id makes the document non-conformant.

The parts also contain separator, continuation-separator, and continuation-notice entries. Those are layout instructions rather than document notes and are not shown in the note list.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§11.3.4, 11.3.7, 17.11.7, 17.11.14, and Part 2, §§6.4, 6.5.2.3.

## Reference order

Visible labels follow first-reference order within the main story rather than note-part order or raw ids, which commonly begin after reserved separator ids. Exact numbering formats, custom marks, and section restart rules are preserved in OOXML but are not reproduced by the screen label. A reference carrying `customMarkFollows` does not draw an automatic inline label.

## Editing boundary

Footnotes and endnotes are separate document stories rather than text owned by their main-story reference nodes. Their current plain-text projection therefore remains read-only. Editing support must model each note body as an editable story, preserve the corresponding note-part structures, and integrate its placement with document layout; rewriting reference metadata or exposing the projection as an isolated text field is not a document-editing model.
