---
"@portone/docx-editor": patch
---

Comment-only verification and export now share the definitions for comment package parts. Public types are unchanged.

Verification rejects new or altered content around comment entries, including outside the XML root, while accepting annotations preserved from the original and those removed by normal comment-part rewrites. Namespace rebindings under rewritten comment markup are also rejected.
