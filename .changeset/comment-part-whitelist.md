---
"@portone/docx-editor": patch
---

Stop `onlyCommentsChangedBy` from excusing a part a submission relates as a comment part.

The three comment parts are left out of the byte comparison, and which parts those were came from the submitted file's own relationships.
A file could relate a second comments, extended comments or people part at any part it liked and have that part go uncompared, so a rewritten styles part, a settings part pointing at a template off the package, replaced image bytes or forged document properties were all reported as a change to nothing but comments.
The parts left out are now the ones the reader opened, a comment part may be related once, one related for the first time has to be a part the submission brought with it, and a relationship part naming one id twice is turned down.

A file that relates an extended comments part but carries no comments part now has that part read, so writing the first comment into it writes that part rather than a second one beside it.
Two extended parts related at once was a file whose settled threads a reader would lose, since only the first of them is read.

This affects 0.2.0. A server that accepted files on this verdict should upgrade and run the check again over what it accepted, where such a file now answers `part-changed` naming the part, or `relationship-changed` naming the relationship part.
