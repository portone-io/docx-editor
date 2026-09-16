---
"@portone/docx-editor": patch
---

Open a table whose row stands inside a content control, instead of standing the whole table down.

A `w:sdt` around a `w:tr` used to be markup nothing read, so one such control cost every row, cell, merge and paragraph of the table it stood in: the table opened as an uneditable placeholder. The control is now read onto the row it wraps, exactly as a control around a cell has always been read onto the cell, so the table is editable and the wrapper goes back around the same row on export. A table nobody edited still goes out byte for byte, and an edit in another row rewrites the table with the control back where it stood.

The lock such a control carries is honored on the row: typing anywhere in a locked row is refused, so is a paragraph or formatting command inside it and a change to its height, while a row beside it is edited as ever. A row deletion is the control being taken away whole, which a lock against deletion refuses; a column deletion reaches into what the control holds, which a lock over its contents refuses. A `w:group` shuts a row's contents without a lock, lifting a lock reaches a row, and a row inserted beside a wrapped one carries no control, so no second control claims the first one's `w:id`.

A control the row cannot carry back out still keeps the table whole: one holding more than the single row the specification describes, one holding another control instead of a row, and one around a row that only continues a vertical merge, since such a row is rebuilt on export rather than written from the document.
