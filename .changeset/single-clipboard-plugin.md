---
"@portone/docx-editor": patch
---

Paste over selected table cells now follows the table's own rule: a copied grid of cells fills the grid, and anything else goes into every selected cell instead of dropping its text into the first one. Shift+V pastes plain text, dropping the formatting the clipboard carries. The right-click Paste puts in what Ctrl+V puts in, images included.
