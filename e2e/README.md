# Real-browser tests

This suite covers IME composition, focus, keyboard interaction, scrolling, and layout behavior that jsdom cannot verify.

Add a browser test when the behavior depends on a real layout, browser event pipeline, or focus model. Keep detailed cases in the spec and add one summary row below for every top-level spec file.

## Running the suite

Run `pnpm test:e2e` from the package root. The suite uses a locally installed Chrome, runs one worker, starts the Vite harness automatically, and is not part of `pnpm test`.

IME tests drive Chrome's real composition pipeline through CDP `Input.imeSetComposition`. They cover Japanese kana-kanji conversion, Chinese pinyin conversion, and Korean hangul composition.

## Coverage

| File | Coverage |
| --- | --- |
| `composition.spec.ts` | Japanese and Chinese composition, candidate replacement, and plain input |
| `hangulComposition.spec.ts` | Hangul syllable assembly, batchim movement, deletion, and locked content |
| `keys.spec.ts` | Enter during an active composition |
| `pageLayout.spec.ts` | Stable page measurement during composition |
| `tablePagination.spec.ts` | Long-table row boundaries, repeated headers, and page-gap placement |
| `pageBreak.spec.ts` | Page-break placement and editing behavior |
| `lockedContent.spec.ts` | Composition at and inside locked content boundaries |
| `keyboard.spec.ts` | Keyboard-only menus, toolbar navigation, focus return, and link-panel focus |
| `linkCard.spec.ts` | Link-card placement, stability, tab access, and focus return |
| `comments.spec.ts` | Comment composition, anchoring, thread actions, scrolling, and narrow layouts |
| `commentMode.spec.ts` | Commenter commenting, refused typing, and right-click menus by mode |
| `clipboard.spec.ts` | Rich HTML formatting and web images through the browser paste pipeline |
| `notesAndPages.spec.ts` | Notes alignment, displayed page numbering, and the demo's closing order |
| `responsiveLayout.spec.ts` | Stable narrow-screen pagination and horizontal scroll, responsive comments, and viewport-contained popovers |
| `tableEditing.spec.ts` | Table-separator editing, pointer-based row resizing, and stable page-boundary previews |
| `tabs.spec.ts` | Tab insertion, formatting, caret and selection geometry, custom alignment, wrapping, links, and RTL layout |

## Layout

| Path | Role |
| --- | --- |
| `../playwright.config.ts` | Chrome project, single worker, and harness server |
| `harness/` | Vite page that mounts a fixture and exposes a test surface |
| `support/harness.ts` | Accessors for document text, block XML, and page state |
| `support/ime.ts` | CDP composition driver |

## Known limits

- Chrome stalls when `Input.dispatchKeyEvent` is interleaved with `Input.imeSetComposition`, so swallowed IME key events and hangul Backspace keys are represented through composition-buffer changes instead.
- CDP does not produce a `contextmenu` event for the context-menu key or Shift+F10, so the harness raises the same browser event before the test follows the real menu path.
- Chrome loses some native selection extensions when the main thread is busy, so a run of Shift+Arrow presses selects a shorter stretch than it asked for and sometimes nothing at all. A selection a test needs as a precondition is set through `selectText`, which dispatches it in one transaction and then waits for the editor to hold both that selection and the focus.
- Canceling a composition with Escape is not covered yet.
- The suite supplies composed IME buffers rather than deriving them from operating-system keystrokes. It verifies the editor's handling of composition, not the operating system's IME implementation.
