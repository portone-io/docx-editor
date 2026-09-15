---
name: docx-review
description: Code review for this repository with three reviewers - structure, behavior, documentation - whose findings are merged, verified against the code, and shown to the user before anything is changed. Use for "review this", "look before I open the PR", or "/docx-review [PR number | branch | path] [--defer <list>]".
user-invocable: true
---

# docx-review

Three reviewers, each owning one set of perspectives, then one merged list for the user to choose from.
The perspectives below are this skill's own; each names the document it is derived from, so a rule that changes there changes here.

## 1. Pick the target

| Argument | Diff |
|---|---|
| none | `origin/main..HEAD` plus uncommitted working-tree changes |
| PR number | `gh pr diff <number>` |
| branch | `origin/main..<branch>` |
| path | that path's `origin/main..HEAD` diff |

`--defer <list>` names work the user has already assigned to a later PR (for example "locks, pagination, site docs").
Pass it to every reviewer so a finding in that area is reported as deferred, not as a defect.

Run `git fetch origin main` first.

## 2. Spawn exactly three reviewers in one message

`model: "opus"`, `subagent_type: "general-purpose"`, names `review-structure`, `review-behavior`, `review-docs`.
They are independent, so spawn them in parallel and spawn no more.

| Reviewer | Perspectives | Derived from |
|---|---|---|
| `review-structure` | One place per concern: a new case should cost one edit, not one per level or dispatcher. Reuse of the helpers a module already has before adding a sibling. Names that say what a thing is in the project's vocabulary. Folder dependency direction. | `docs/architecture.md`, `src/folderBoundaries.test.ts`, the naming rule in `CLAUDE.md` |
| `review-behavior` | Defects and regressions, including a behavior the change made reachable that used to be unreachable. Round trip: untouched content goes out byte for byte, an edited export re-reads to the same shape, exports validate against the schema. Every consumer of the document model - locks, guards, clipboard, identities, pagination, sections, headers and footers - handles a new node or attr. Tests that pin the behavior rather than the implementation. | `docs/testing.md` (the rule-guarding tests), `spec/notes/`, `__fixtures__/README.md` |
| `review-docs` | Comments that restate a name, a type, or control flow, or run longer than the surrounding text. Spec notes that record more than the decision and its ground. Site pages that still claim the old behavior, or gained a section where a sentence would do. Changeset present, at the level the change warrants, written for release notes. | `CLAUDE.md` (Comments, Context routing), `site/AGENTS.md`, `CONTRIBUTING.md` (Changesets) |

Every prompt carries:

- The working directory, the diff range, and the deferred list. No browser or screenshot checks.
- Read `CLAUDE.md` first, then only the documents your row names. The other rows are someone else's.
- Verify every candidate against the code: follow the call path, the types, the tests. Where the claim is about runtime behavior, prove it with a throwaway probe test run through `pnpm vitest run <file>` and delete the file afterwards. Drop what you cannot confirm.
- Change no project file. Do not commit.
- Write the result to `<scratchpad>/review-<name>.md` and reply with the path and the count only; a long reply is truncated.
- The file holds one JSON array: `[{ file, line, summary, failure_scenario, confidence: "confirmed" | "plausible", deferred: boolean }]`, most severe first, empty when nothing is found. `summary` is the defect in one sentence; `failure_scenario` names the input or state and what goes wrong.

## 3. Merge and show

Read the three files, then TaskStop the three reviewers at once.

- Merge duplicates into one finding.
- Order: defects and regressions, then round-trip and specification deviations, then structure and naming, then comments and documentation.
- Set aside: a structure that predates the change (list it under "outside this change"), and anything in the deferred list (list it under "deferred", so the later PR's author sees it).
- Show the user three groups: apply, drop, outside this change or deferred.
- Stop here. Nothing is applied until the user picks.

## 4. Apply

Delegate only the findings the user picked to one `model: "opus"` agent, each with the direction of the fix.

- The agent's rules: a new comment only for a constraint the code cannot state, one line; no `any`, no `as`; no commit; while iterating run only the test files that cover the change, then `pnpm check` once at the end.
- When the report arrives, re-read every new comment line in the diff, then TaskStop the agent.
