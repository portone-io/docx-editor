# Headers and footers

## Parts and section references

Header and Footer parts are related from the Main Document part. A section selects them with `w:headerReference` and `w:footerReference`; each `r:id` must identify an internal relationship of the matching type. The editor reads every part a section names as a story of its own and resolves each section's references against them, so a second section shows the parts it names rather than the first section's. A related part no section names is a leftover rather than content, so it is never read: its bytes are repacked as they came, and markup no consumer would draw is no reason to refuse the file. A part whose story nobody rewrote is repacked unchanged; a rewritten one is written back around the two ends it arrived with, its prolog included.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§11.3.6, 11.3.9, 17.10.2, 17.10.5.

## Variant selection and page fields

The `default` story applies to odd pages and to all pages when odd/even stories are disabled. `w:titlePg` selects the `first` story for the section's first page, and `w:evenAndOddHeaders` in Settings enables the `even` story. When an enabled first or even variant is not declared in the section being drawn, its displayed story is blank rather than falling back to `default`.

The on-screen page number begins at `w:pgNumType/@w:start` when declared and at 1 otherwise. That number determines odd/even selection and replaces `PAGE`; the visual page count replaces `NUMPAGES`. Both simple and complex forms of these two fields are recognized and projected as decimal numbers. Each section applies its own start to the visual page number; restarting the count at a section boundary is not reproduced, so a later section's numbers carry on from the page's place in the whole document. Other `w:pgNumType` settings, including number formats and chapter numbering, remain preserved but are not reproduced. Their source XML and cached results remain untouched.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§17.10.1–17.10.6, 17.16.5.29, 17.16.5.45.
