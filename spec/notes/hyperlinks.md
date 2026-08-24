# Hyperlinks (`w:hyperlink`)

## Two kinds, and which wins

`w:hyperlink` (§17.16.22) wraps runs the way a content control does, and names its target one of two ways.
`r:id` points at a relationship in the document part's rels, whose Target is the address; `w:anchor` names a bookmark in the same document.
When both are given, `r:id` shall supersede the anchor.
An `anchor` naming no existing bookmark falls back to the start of the document, as does an omitted one.
`w:docLocation` and `w:history` ride along; neither changes what the link points at.

The address of an external link therefore never appears in the body XML.
Writing a new link means writing a relationship, the same move inserting an image already makes (`docx/media`), with one difference: a hyperlink relationship carries `TargetMode="External"`, which the image writer never sets because a media target is a part inside the package.

Observed 2026-08-20 against ECMA-376 5th edition, Part 1.

## Target preservation

External links resolve through `r:id`; anchor-only links have no external address to expose. The wrapper's original opening XML retains properties that do not participate in target resolution, including `w:history`, `w:docLocation`, and `w:tooltip`. Splitting an external link may leave both halves pointing at one relationship because OPC relationships are not single-use.

## Where the prose and the package part company

The example under `r:id` writes the relationship as `<Relationship Id="rId9" Mode="External" Target="…"/>`.
There is no `Mode` attribute: the relationship element is OPC's (ECMA-376 Part 2), and the attribute that says the target lies outside the package is `TargetMode`, which is what Word writes and what we write.
Read the example as prose about intent rather than as markup to copy.
