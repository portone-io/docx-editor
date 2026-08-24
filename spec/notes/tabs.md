# Tabs

ECMA-376 Part 1 §17.3.3.32 places a tab at the next custom stop beyond the preceding laid-out content, then at the next multiple of `defaultTabStop` when no later custom stop exists. Section §17.15.1.25 supplies a 720-twip interval when the document omits that setting.

Custom stops are additive across paragraph-property layers (§17.3.1.38). The display resolves document defaults, numbering-level paragraph properties, paragraph styles, and direct paragraph properties. A `clear` entry removes a stop inherited at the same position. The display creates the implicit stop associated with a hanging indent; it does not currently suppress that stop when the preserved `w:noTabHangInd` compatibility setting is present. Table-style-only paragraph properties remain in their source XML but are outside the current paragraph display model.

The display model normalizes the Transitional `left` and `right` values to the logical `start` and `end` values. A `bar` entry is a drawing instruction under §17.18.84 and is therefore excluded when choosing the next positioning stop. The legacy `num` value is retained in the model but positioned as `start`, matching its compatibility role between a list label and paragraph content. Leader values remain part of the effective stop, while the display deliberately does not synthesize leader glyphs or bar lines. None of these display decisions replaces the preserved paragraph-property XML used for export.

When a `center`, `end`, or `decimal` stop would require its following text to overlap content already laid out on the line, the display keeps the following text at the current position. The selected custom stop and its source XML remain unchanged.
