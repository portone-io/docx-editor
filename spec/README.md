# OOXML specification workflow

This project uses ECMA-376 5th edition for OOXML format and conformance decisions. A section citation in code names its part when it is not from Part 1.

## Working with the specification

1. Check the [existing specification notes](./notes/README.md) for a decision already made.
2. If no note covers the question, read the relevant ECMA-376 part and section rather than relying on memory.
3. Use the specification to determine document format and conformance requirements. Treat editor interactions as project decisions unless those requirements constrain them.
4. Record a reusable decision under `notes/`, citing the relevant part and section without copying the specification prose.

The standard does not need to be consulted for changes that do not interpret or write OOXML.

## Local material

`schemas/transitional/` contains the 26 W3C XML Schemas distributed with ECMA-376 Part 4. They are committed unchanged for export validation and must not be edited. [Third-party notices](../THIRD_PARTY_NOTICES.md) records their source, hash, copyright notice, license, and disclaimer.

The specification PDFs are downloaded into the gitignored `pdf/` directory when needed:

```sh
scripts/fetch-ooxml-spec.sh
```

The script downloads and verifies Parts 1 and 4. Part 1 is the main markup reference, while Part 4 defines Transitional features used by documents produced by common office applications.

## Finding a section

`pnpm spec` builds a local index over the downloaded PDFs and looks up a section number or title:

```sh
pnpm spec
pnpm spec 17.5.2.23
pnpm spec lock
```

The generated index remains under `pdf/` with the PDFs and is not committed.

## Sources

| Item | Source | SHA-256 |
| --- | --- | --- |
| Part 1, 5th edition (2016) | https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip | `9d0bcad9cf06054785b03762fcfadbf6bab7e54a5f9d69434e34b7fd464d4129` |
| Part 4, 5th edition (2016) | https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip | `bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd` |
