---
"@portone/docx-editor": patch
---

Keep a picture's extension lists intact when the image is resized.

A picture written by Word often carries extension lists - the small records a word processor
attaches to a picture to remember things the format has no field for, such as which application
created it or what resolution to keep it at. Resizing such an image used to overwrite the name of
any empty record with the image's new size, which left the document with content Word reads as
invalid. Only the two sizes an image is drawn from are rewritten now, and every extension record
goes back out exactly as it came in.

Size updates follow the picture’s own XML paths, so an extension’s nested transforms and size
records stay unchanged too. Size attributes retain their quoting and surrounding markup.
