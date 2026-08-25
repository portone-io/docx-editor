# Security policy

## Reporting a vulnerability

Report a vulnerability through [GitHub's private advisory form](https://github.com/portone-io/docx-editor/security/advisories/new). The report stays private to the maintainers until a fix is released.

Do not open a public issue for a vulnerability, and do not include a document holding confidential content in the report.

A report is most useful with the version it affects, what an attacker gains, and the smallest input that reproduces it. A DOCX file that triggers the problem is worth attaching once it has been stripped of anything confidential.

We aim to acknowledge a report within three working days and to say what we intend to do about it within ten.

## What is in scope

This is a library that runs in the browser as part of another application. A finding is in scope when this package's own code is what makes it exploitable:

- a crafted DOCX file that, when imported, escapes the editor's document model, reaches the surrounding page, or runs script;
- an exported document that carries content it was never given;
- content from a document reaching the DOM without being treated as untrusted.

Out of scope: vulnerabilities in the browser, in an application embedding this editor, or in a dependency where this package's own use of it is not what exposes the problem. Report those to the project that owns them.

## Supported versions

Fixes go to the latest release. This package is before 1.0, so a fix may arrive in a minor version rather than a patch.
