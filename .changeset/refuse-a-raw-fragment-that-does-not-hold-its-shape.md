---
"@portone/docx-editor": patch
---

Turn down a raw OOXML fragment that does not hold its shape as it enters the document, rather than writing it into the exported file.

A paragraph, a run, a table, an image, an annotation and anything import could not model all carry their original XML, and the editor draws it into the page as a `data-` attribute so that it can read the live DOM back after an IME composition or a browser edit. Nothing checked what came back. A fragment written by a consumer plugin, or handed to `view.pasteHTML`, could close the element it was about to be spliced into and open another: a `data-ppr` reading `</w:p><w:p><w:pPr>...` wrote a second paragraph into the exported body, and one that never closed reached export and failed the whole document with `malformed-xml` without naming what broke it.

Every rule that reads raw XML now holds it to the shape that attr goes back out as: a whole element under an expected name, an opening tag's attributes, or an opening tag the writer closes itself. A fragment that does not hold it is turned down and the rule gives up, so the content settles one level plainer instead - a paragraph loses its properties and keeps its text, a run loses its mark, a content control or a hyperlink loses its wrapper. What import could not model still travels as the element it is, whatever namespace that is in, and every fragment the editor draws is read back unchanged.
