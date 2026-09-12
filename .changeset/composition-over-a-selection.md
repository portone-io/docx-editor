---
"@portone/docx-editor": patch
---

Replace the selected text when an IME composition is typed over it. A selection beginning at a preserved marker - a bookmark, a note reference, the number a footnote is drawn by - was left standing by the browser, so the composed syllables arrived in front of the words they were meant to replace, or nothing happened at all.
