---
"@portone/docx-editor": patch
---

Draw nothing for a content control a paragraph holds with nothing inside it.

A `w:sdt` standing between the words of a paragraph whose `w:sdtContent` was empty, or which wrote no content element at all, used to open as a dotted box saying the content was preserved. A server that answers a failed condition by leaving a tagged phrase empty - a clause that does not apply, a signing date to be filled in later - therefore drew a box in the middle of a sentence, in every draft, and the reader could neither open it nor take it away. Such a control now opens as a node of its own that holds nothing and takes no width, so the sentence reads on as the file says it should.

Everything the control states about itself is kept: its tag, its id, its lock and its data binding all ride back out untouched, a control nobody edited goes out as the bytes it arrived as, and a control standing inside another control or inside a hyperlink still has that wrapper closed around it. Backspace and Delete beside it pass over it and take the character on the far side, as if it were not there, so the control goes only with a selection that covers it, and then only if its lock does not say otherwise; a control holding nothing that carries a lock also keeps a selection covering it from being replaced or deleted. Copying a sentence it stands in leaves the control behind, since it holds nothing to carry, and a document holding one no longer reports it as content that was preserved rather than read, since nothing of it was lost.

The control a file writes with nothing inside it between paragraphs, which 0.6.3 released as one that goes only when it is selected whole, now follows the same rule: it can no longer be selected on its own, and a selection covering it is what removes it.
