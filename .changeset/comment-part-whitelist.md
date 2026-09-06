---
"@portone/docx-editor": patch
---

Stop `onlyCommentsChangedBy` from excusing a part a submission relates as a comment part.

The three comment parts are left out of the byte comparison, and which parts those were came from the submitted file's own relationships.
A file could relate a second comments, extended comments or people part at any part it liked and have that part go uncompared, so a rewritten styles part, a settings part pointing at a template off the package, replaced image bytes or forged document properties were all reported as a change to nothing but comments.
The parts left out are now the ones the reader opened, a comment part may be related once, and one related for the first time has to be a part the submission brought with it.

This affects 0.2.0. A server that accepted files on this verdict should upgrade and run the check again over what it accepted, where such a file now answers `part-changed` naming the part, or `relationship-changed` naming the relationship part.
