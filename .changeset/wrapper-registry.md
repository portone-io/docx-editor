---
"@portone/docx-editor": patch
---

Hyperlinks that hold content controls, and nested controls, open as editable text.

Only one of the arrangements used to be readable - a control holding a link. A link holding a
control, or a control holding a control, brought the inner one in as a small box naming the element
it stood for, and the text it held could not be typed in. Any depth and any order of the two is now
ordinary text wearing a wrapper each, and goes back out nested as it came. A hyperlink inside a
hyperlink is the one arrangement still kept whole, since a link marks the text it covers once.

A link made on text inside a control is written inside that control, and a lock put on a stretch
that sits wholly inside a link is written inside that link, instead of either wrapper being split
around the other. Where a control holds a control, locking shuts the outer one and lifting a lock
opens the inner one the selection stands in.

Plugins reading the drawn page will find `data-key` and `data-depth` on a control and a link in
place of `data-sdt-key` and `data-link-key`.
