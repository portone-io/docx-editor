---
"@portone/docx-editor": patch
---

The type declarations now resolve in projects that use TypeScript's `node16` or `nodenext` module resolution.
Their relative imports left the file extension off, which those settings do not accept, so such a project got `any` for almost every export, or errors inside the package with `skipLibCheck` off.
Heads-up: code in such a project that was going unchecked is now typechecked against the package and may report type errors it had all along.
