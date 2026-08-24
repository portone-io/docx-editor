# Features

`@portone/docx-editor` edits common DOCX body content while preserving document structures outside its editing surface.

This page is the user-facing support summary. Add a claim only when implemented behavior and tests support it, classify it as editable, preserved, or limited, and describe the most important constraint once. Keep API details in [Custom controls](./custom-controls.md) or [programmatic DOCX import and export](./core.md), and OOXML rationale in [specification notes](../spec/notes/README.md).

## Supported features

| Area | What you can do |
| --- | --- |
| Text formatting | Edit bold, italic, strikethrough, underline, palette or custom HEX colors, highlight, fonts, and font size. |
| Paragraphs | Apply paragraph styles, alignment, indentation, line spacing, line breaks, and page breaks. |
| Tabs | Insert, paste, format, navigate, select, and align document tabs using the document's automatic interval and effective paragraph stops. |
| Lists | Create and edit common bullet and numbered lists, including nested levels. |
| Tables | Insert and remove rows, columns, and tables; merge, split, resize columns and rows, format, vertically align, and pad cells. Long tables continue at safe row boundaries, repeating leading `tblHeader` rows. |
| Images | Insert, paste, drop, select, and resize inline images. |
| Hyperlinks | Create, change, remove, and open external links. |
| Comments | View anchored threads; create, edit, reply to, resolve, reopen, and delete plain-text comments. |
| Footnotes and endnotes | View references and plain-text note bodies while editing surrounding text. |
| Headers and footers | Visually preview the first section's plain-text first, even, and default stories, including decimal `PAGE` and `NUMPAGES` fields. |
| Content controls | Edit or lock supported text and table-cell controls. |
| Clipboard | Paste supported character formatting, paragraphs, destination-matched heading styles, line breaks, links, and common lists. Unsupported web styling is reduced to readable text. |
| Editor | Use built-in controls, selectable document zoom, read-only mode, undo and redo, keyboard navigation, and IME composition. |

## Preserved content

| Content | Behavior |
| --- | --- |
| Additional Word character formatting and character styles | Existing formatting that the editor does not display or edit is preserved. |
| Bookmarks | Existing bookmark ranges are preserved at their document positions but are not displayed or editable. |
| Document fields | Existing fields in the document body are preserved but not displayed or editable. |
| Tracked changes | Existing revision markup is preserved but not displayed or editable. |

## Current limitations

- Uncommon list formats may be displayed as decimal numbers, and a new list requires an existing numbering part.
- Tab leader glyphs and bar lines remain preserved in OOXML but are not drawn.
- Visual page boundaries are approximate and may differ from Word or printed output.

## Preservation

Content outside the supported editing surface is not silently discarded. Untouched package parts and document blocks are written back from their original XML, while placeholders retain structures the editor cannot safely model.

If a safe round trip cannot be guaranteed, import or export fails with a stable error code instead of returning a damaged document. [Architecture](./architecture.md#preservation-model) explains the preservation model, and [programmatic DOCX import and export](./core.md) documents package limits and server parser requirements.
