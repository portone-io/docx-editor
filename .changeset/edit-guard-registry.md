---
"@portone/docx-editor": patch
---

Commands and `canRunCommand` now report false where a bookmark marker or a note reference would be removed, instead of reporting true and changing nothing.

Both are preserved rather than edited, and the rule that held them stood outside the guard every command asks, so a control drawn from a command stayed live over a marker and swallowed the click.
Every rule an edit is judged by is now registered in one list, which the runtime filter, the dry run behind `canRunCommand`, and the command predicates all walk, so a refusal over a marker also ends an open IME composition the way a refusal over a lock already did.
