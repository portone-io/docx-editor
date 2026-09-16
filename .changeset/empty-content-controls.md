---
"@portone/docx-editor": patch
---

Draw nothing for a content control with nothing inside it.

A `w:sdt` whose `w:sdtContent` was empty, or which wrote no content element at all, used to open as a grey box saying the content was preserved. A server that answers a failed condition by leaving a tagged control empty therefore drew a box for the clause that is not there, and the reader could neither open it nor take it away. Such a control now opens as a node of its own that holds nothing and takes no room on the page, so the clause reads as the nothing the file says it is.

Everything the control states about itself is kept: its tag, its id, its lock and its data binding all ride back out untouched, and a control nobody edited goes out as the bytes it arrived as. Backspace and Delete beside it pass over it, so it goes only when it is selected whole, and then only if its lock does not say otherwise; a copy of one gets an id of its own the way every other control does.
