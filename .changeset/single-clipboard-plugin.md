---
"@portone/docx-editor": patch
---

Pasting text over selected table cells now fills every selected cell instead of only the first. Pasting the clipboard's plain text (Ctrl/Cmd+Shift+V) ignores accompanying HTML and images. Right-click Paste passes the clipboard's HTML and text through the same handlers as Ctrl/Cmd+V, including image loading and text fallback. Unreadable clipboard content leaves the selection intact when there is no usable text fallback.
