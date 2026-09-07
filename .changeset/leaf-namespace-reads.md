---
"@portone/docx-editor": patch
---

Importing `emuToPx`, `pxToEmu` or `toImageExtent` off `./core`, or the image file helpers off `./commands`, no longer carries the XML naming layer into a consumer's bundle. The picture module read the namespace table at the top of the file, which a bundler keeps as a side effect, so one multiplication cost 817 bytes minified where it costs 182. Nothing written into a document changes.
