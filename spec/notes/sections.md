# Sections

## Where a section's properties are written

A document is a sequence of sections, and every section carries its own `w:sectPr`. The last section's stands as the final child of `w:body`; every earlier one stands inside the `w:pPr` of the paragraph that ends that section, so a paragraph-level `w:sectPr` marks the end of a section rather than its start. A body holding no `w:sectPr` at all is still one section, with everything left unstated.

Both the paragraph-level and body-level forms use `CT_SectPr`, so one reader answers for both. `CT_SectPrBase` is the restricted form used inside `w:sectPrChange` for historical section properties; it omits header/footer references and `sectPrChange`.

Observed 2026-09-08 against ECMA-376 5th edition, Part 1, §§17.2.2, 17.6.17, 17.6.18, and `wml.xsd` (`CT_SectPr`, `CT_SectPrBase`).

## Where a break written inside a content control ends its section

A block-level content control holds what the body holds (`CT_SdtContentBlock`, §17.5.2.34), so a paragraph inside one is one of the body's own paragraphs and the `w:sectPr` it carries ends a section exactly as a loose paragraph's does; a paragraph inside a table cell is not part of that sequence, so a `w:sectPr` written in a cell is preserved as markup rather than read as a section.
Word parts the control itself where a paragraph inside one ends a section. This editor lays a control out as one block on one section's paper, so a section ends where the top-level block holding its last paragraph ends and a break standing further up inside a control still takes effect at the control's end.
One block therefore ends at most one section: where a control holds more than one break the first is the one read, and the breaks after it stay preserved markup that goes back out on export. Two sections ending in the same place would leave the later one covering no position at all, which nothing asking what section a position falls in could ever reach.
The guard that refuses to let a section-ending paragraph disappear answers for every `w:sectPr` a paragraph carries, the ones this reading passes over included, so no break is lost without the author saying so.

Observed 2026-09-16 against ECMA-376 5th edition, Part 1, §§17.5.2.34, 17.6.18, and `wml.xsd` (`CT_SdtContentBlock`).

## Splitting and merging the paragraph that ends a section

Because the properties belong to the paragraph that *ends* the section, splitting that paragraph leaves the section break with the later half: the earlier half becomes an ordinary paragraph inside the same section, and the later half goes on ending it. Copying such a paragraph elsewhere carries no break at all, since the copy ends no section.

Merging the two back together is the same rule read backwards, and this editor does not perform it: a transaction that would leave fewer section-ending paragraphs than the document had is refused instead, so a section cannot disappear without the author saying so.

Observed 2026-09-08 against ECMA-376 5th edition, Part 1, §17.6.17.

## Child order and repeated children

`CT_SectPr` is a sequence: the header and footer references first, then the section's own content, then `w:sectPrChange`. The references are the only repeatable children, at most one per `w:type` value (`default`, `first`, `even`), so a reference is identified by the variant it names rather than by its element name alone. `w:type` is required by `CT_HdrFtrRef`; this editor tolerates a missing value by reading it as `default`. The order the registry holds is in [Child order](./childOrder.md).

Observed 2026-09-08 against ECMA-376 5th edition, Part 1, §§17.6.12, 17.6.18, and `wml.xsd` (`EG_HdrFtrReferences`, `EG_SectPrContents`).

## What the editor draws

The screen paginates each block on the paper of the section that block sits in: the page a block opens is as tall as that section's body, and headers, footers and the page-number start are resolved from the same section. Width is the exception, because one sheet is drawn at one width: every page is drawn at the first section's, so a section on wider paper is paginated on its own height and drawn on the sheet's paper. A table is fitted to the body width of its own section and drawn shrunk to fit the sheet when that is wider, and a pasted image is fitted to the width the sheet draws and to the height its own section leaves. A document carries every section's properties back out untouched whatever is drawn. `site/content/docs/features/pages-and-sections.mdx` records for users what is drawn and what is not.

`w:orient` is not applied on top of `w:pgSz`: a producer writes the width and the height already swapped for a landscape section, so honouring both would turn the paper back upright. `w:cols` is preserved in the section XML without a separate model field and is not drawn. The page-start rule in `w:type` is read into the model, and `continuous` is drawn: such a section carries on down the page the section before it ends on, and a keep with next holds across that boundary. A page two sections share that way draws the header and footer of the section it opens with, where Word draws the last section's. Every other kind is drawn as `nextPage`, so `evenPage`, `oddPage` and `nextColumn` each open the very next page rather than skipping to one of the parity they name.
