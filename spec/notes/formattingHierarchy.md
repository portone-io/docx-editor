# Formatting hierarchy

ECMA-376 Part 1 §17.7.2 fixes one order in which the sources of a paragraph's and a run's formatting are applied, and every display value this editor derives is resolved in that order by one function, `resolveParagraph` / `resolveRun` in `src/docx/formatting/resolve.ts`.
Opening a document, applying a style, editing a paragraph, pasting, and formatting the paragraph an edit built all read the same resolver, so a fragment resolves to the same values whichever path asked.

## The order

Lowest first: document defaults (`w:docDefaults`), the table style, the numbering level's paragraph properties, the paragraph style (with its `basedOn` chain already folded, §17.7.1), the character style (`w:rStyle`, §17.3.2.29), then the direct `w:pPr` and `w:rPr`.
A paragraph naming no `w:pStyle` wears the paragraph style marked `w:default="1"`, and a paragraph naming one wears that style instead of it rather than on top of it.

§17.7.2 is not consistent with itself about where the numbering level sits: its prose applies the numbered item's properties before the paragraph style, while its figure draws the paragraph style before numbering.
This editor follows the prose, so a stop or a clear a paragraph style writes beats one the level writes at the same position.

A `w:numId` of 0 at any layer removes the numbering a lower layer gave (§17.9.18); it is not a reference to a list.

## Editor decisions

- The default font and size stay in the sheet's CSS variables so the toolbar can distinguish an inherited document size from a style or direct setting. Other supported `w:rPrDefault` values are folded into display attrs; CSS variables do not carry bold, italic, or other character properties. Removing a direct setting resolves the remaining layers again, including character styles.
- A toggle property (§17.7.3) a run switches off outright (`w:val="0"`) is read as `false`, not as absent, and is drawn as off. Absent means the run says nothing and the layer below stands; false is the run beating that layer. Word draws it the same way.
- Switching a run property off writes the off state into the run (`w:b w:val="0"`, `w:u w:val="none"`, `w:color w:val="auto"`, `w:shd w:fill="auto"`) when a layer below the run switches that property on, and removes the element otherwise. Removing the element where a style switched the property on would let the style win again.
- The indent a numbering level lays down is drawn as a decoration over the paragraph and is not part of its display values, so a list paragraph's own indent can still be told from the inherited one. The implicit hanging-indent tab stop reads the level's indent the same way the decoration draws it: only where the paragraph writes none of its own.
- A numbering level's `w:rPr` dresses the number it draws and never the text of the paragraph that number stands in front of (§17.9.24), so it reaches the marker's own drawing rather than the display values of the paragraph or of its runs. It is not a layer of the hierarchy above: the level's `w:pPr` is (§17.7.2), and its `w:rPr` is not.
- A run holding no characters - nothing but its `w:rPr`, or empty `w:t` (§17.3.2.25, §17.3.3.31) - is read as a node of its own that keeps its properties, rather than as preserved XML with nothing to format. Its highlight or shading (§17.3.2.15, §17.3.2.32) paints behind the run's contents, and an empty run has none, so Word draws nothing there. The editor draws one whose own `w:rPr` highlights or shades it over a short width instead, so a blank a form leaves to be filled in can be seen and pressed; a highlight or shading it only inherits from a style draws nothing, as in Word.
- Text typed at a caret takes the formatting of the nearest character its paragraph shows, the one before the caret or else the one after it, which is how Word formats typed text. What draws nothing is passed over: a bookmark, a comment's range marker, a control holding nothing, and a run holding no characters unless the editor draws it as a blank. A drawn blank counts as a character, and what is written where it gives the formatting takes its properties and replaces it, a paste included. In a paragraph that shows nothing, a run holding no characters gives the formatting, because Word's source there, the paragraph mark's run properties (§17.3.1.29), is not read by this editor.
- A table style contributes what it lays down for the whole table and then what it dresses each part the cell belongs to with, in the order §17.7.6 gives; [Table styles](./tableStyles.md) records how the parts are decided.

Tab stops are additive across the layers (§17.3.1.38); [Tabs](./tabs.md) records how the stops are resolved and which layer each effective stop is attributed to.
