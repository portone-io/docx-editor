---
"@portone/docx-editor": patch
---

Say which content blocks an export, and where it stands.

Every entry `exportProblems`, `documentExportProblems` and a `blocked` `downloadDocx` report now carries a `reason`, a discriminated union naming the situation and the content it is about: `unwritten-story-change` with the story and what was done to it, `undefined-list` with its `numId`, `duplicate-preserved-block` with the node and the story it stands in, `unmatched-bookmark`, `missing-content-types` with the part, and the rest. It tells apart the situations one code covers, so a host can write the sentence its own user reads instead of matching an English message. The codes and messages are unchanged, so a `switch` over `code` keeps working.

A problem about a footnote or an endnote now carries `pos` as well, the first reference to that note in the body, since the note's own text stands nowhere there. A note nothing refers to still has no position.

`DocxExportError` carries the entry it was raised for as `problem`, so a `catch` reads the same `reason` and `pos` without asking again. A refusal met while writing that the invariants did not predict leaves it absent.

Since 0.6.1, `exportBytes` and `exportDocx` refuse a document holding a change nothing writes back - a header or footer story added or removed, an edited footnote or endnote separator - instead of saving the file without it. Check `exportProblems` before writing, or catch `DocxExportError`.
