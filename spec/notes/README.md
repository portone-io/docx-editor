# Specification notes

These notes record reusable OOXML interpretations, interoperability findings, and project decisions that affect how a document is read, written, or preserved. They are not a feature matrix or an implementation diary.

- [Comments](./comments.md)
- [Content controls](./contentControls.md)
- [Table cell layout](./tableCellLayout.md)
- [Headers and footers](./headersAndFooters.md)
- [Footnotes and endnotes](./footnotesAndEndnotes.md)
- [Hyperlinks](./hyperlinks.md)
- [Bookmarks](./bookmarks.md)
- [Tabs](./tabs.md)

Do not add a note merely because an implementation touches OOXML. Keep local behavior in code and tests, user-visible support in `docs/features.md`, and API contracts in the API documentation. Add one file per OOXML area only when a standards interpretation or interoperability decision is reusable across features and would be difficult to recover from those sources. Update an existing note when possible so this directory does not become another implementation index.

Distinguish format requirements from editor choices, omit implementation details that are already clear from code, and follow the [specification workflow](../README.md#working-with-the-specification). Cite the relevant part and section or another authoritative source for each finding.
