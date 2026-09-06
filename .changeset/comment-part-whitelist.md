---
"@portone/docx-editor": patch
---

Stop `onlyCommentsChangedBy` from excusing a part a submission relates as a comment part.

The three comment parts are left out of the byte comparison, and which parts those were was read from the submitted file's own relationships.
A file could relate a second comments, extended comments or people part pointing at any part it liked and have that part go uncompared, so an edit to the styles, the settings or an image was reported as a comment-only change.
The parts left out are now the ones the reader opened, a comment part may be related once, and one related for the first time has to be a part the submission brought with it.
