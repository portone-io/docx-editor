# Headers and footers

## Parts and section references

Header and Footer parts are related from the Main Document part. A section selects them with `w:headerReference` and `w:footerReference`; each `r:id` must identify an internal relationship of the matching type. The editor reads the first section only and repacks every related part unchanged.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§11.3.6, 11.3.9, 17.10.2, 17.10.5.

## Variant selection and page fields

The `default` story applies to odd pages and to all pages when odd/even stories are disabled. `w:titlePg` selects the `first` story for the section's first page, and `w:evenAndOddHeaders` in Settings enables the `even` story. When an enabled first or even variant is not declared in the first section, its displayed story is blank rather than falling back to `default`.

The on-screen page number begins at `w:pgNumType/@w:start` when declared and at 1 otherwise. That number determines odd/even selection and replaces `PAGE`; the visual page count replaces `NUMPAGES`. Both simple and complex forms of these two fields are recognized and projected as decimal numbers. Other `w:pgNumType` settings, including number formats and chapter numbering, remain preserved but are not reproduced. Their source XML and cached results remain untouched.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§17.10.1–17.10.6, 17.16.5.29, 17.16.5.45.
