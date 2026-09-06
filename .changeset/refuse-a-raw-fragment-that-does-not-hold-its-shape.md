---
"@portone/docx-editor": patch
---

Turn down a raw OOXML fragment that does not hold its shape as it enters the document, rather than writing it into the exported file.

Keep imported formatting when a document uses an inherited default namespace or an alternative WordprocessingML prefix. Reject nested namespace rebinding that could disconnect hyperlinks from their targets on export.

The editor draws the original XML of a paragraph, a run, a table, an image or an annotation into the page as a `data-` attribute, so that it can read the live DOM back after an IME composition or a browser edit. A fragment written into one of those attributes from outside, by a consumer plugin or through `view.pasteHTML`, is now held to what that attribute goes back out as: a whole element under an expected name, an opening tag's attributes, or an opening tag the writer closes itself. One that does not hold it is turned down along with the rule reading it, so the content settles one level plainer - a paragraph keeps its text and loses its properties, a run loses its mark, a content control or a hyperlink loses its wrapper - instead of a `data-ppr` reading `</w:p><w:p>...` writing a second paragraph into the exported body, or one that never closed reaching export and failing the whole document with `malformed-xml`.
