# Child order

WordprocessingML fixes the order of a property element's children in its schema rather than in its prose.
So the spot a child that was not there goes into is read off the complex type the parent element carries, and `spec/schemas/transitional/wml.xsd` is the source for it.

## A sequence fixes one order, a repeating choice fixes none

Most property types are an `xsd:sequence` of optional elements, which admits exactly one order: CT_PPr (`w:pPr`, §17.3.1.26), CT_TcPr (`w:tcPr`, §17.4.69), CT_TblPr (`w:tblPr`, §17.4.59), CT_SdtPr (`w:sdtPr`, §17.5.2.38), and the border and margin types all read that way.
A document writing those children in another order is invalid, and a schema validator says so.

A handful of types instead hold a choice that may come round again, and those fix no order among its members at all.
CT_RPr (`w:rPr`, §17.3.2.28) is EG_RPrBase as a repeating choice, and CT_TrPr (`w:trPr`, §17.4.81) extends CT_TrPrBase, which is a repeating choice of its own.
Any permutation of those children is valid, so a validator will not catch an order that changes, and the order to write them in is the writer's to decide.
CT_SectPr holds a smaller case of the same thing: `w:headerReference` and `w:footerReference` come from a group the type takes up to six times over.

What stays fixed even in those types are the boundaries the surrounding sequence draws.
CT_RPr puts `w:rPrChange` after everything the choice holds, and CT_TrPr puts `w:ins`, `w:del`, and `w:trPrChange` after it.

## Where a choice leaves the order open, this editor writes the schema's own order

The convention chosen for a repeating choice is the order the schema itself lists its members in.
Alternatives that exclude one another - the twelve kinds a content control may declare, the three markup records a cell may carry - are listed in that same order, since only one of them can ever stand there.
Nothing about a document depends on this choice; it is picked so that one order serves every writer and so that a reviewer can check a registered order against the schema line by line.

## A paragraph mark's run properties are not a run's

The `w:rPr` inside a `w:pPr` (§17.3.1.29) is CT_ParaRPr, not CT_RPr.
It puts `w:ins` (§17.13.5.20), `w:del` (§17.13.5.15), `w:moveFrom`, and `w:moveTo` ahead of everything a run's own properties hold, because those record that the paragraph mark itself was inserted, deleted, or moved.
Writing a paragraph mark's properties in a run's order would put revision markup in a spot the schema does not allow, so the two orders are registered apart and the enclosing element decides which applies.

Observed 2026-09-07 against ECMA-376 5th edition, Part 1, and the transitional schemas distributed with Part 4.
