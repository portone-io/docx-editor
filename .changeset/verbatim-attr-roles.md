---
"@portone/docx-editor": patch
---

An untouched table opened in the editor is written back byte for byte, its `w:tblGridChange` included.

Export decided whether a block was untouched by comparing it with the one import produced, attrs and all.
Opening a document works some of those attrs out again from the formatting around them - a table's shared cell borders among them - so a table nobody had touched compared unequal and was rebuilt, and the rebuild dropped the markup the writer does not model.

Each attr now declares whether the writer writes from it, whether the editor works it out for the screen, or whether it identifies something in the open document, and the comparison ignores the ones worked out for the screen.
Exporting through `createEditorState` is held to the same byte identity as exporting straight from `importDocx`.
