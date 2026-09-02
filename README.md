# @portone/docx-editor

A simple DOCX editor for the browser that reads and writes OOXML directly while preserving the document structure around each edit. Edit text, formatting, lists, tables, images, links, and locked content controls without an HTML conversion step.

<p align="center">
  <img src="./assets/editor.png" alt="A DOCX document open in the editor" width="960" />
</p>

## Quick start

```sh
npm install @portone/docx-editor
```

```tsx
import "@portone/docx-editor/styles.css";

import { DocxEditor, type DocxEditorHandle } from "@portone/docx-editor";
import { useRef } from "react";

export function Editor({
  file,
  user,
}: {
  file: File;
  user: { id: string; name: string };
}) {
  const editorRef = useRef<DocxEditorHandle | null>(null);

  return (
    <DocxEditor
      ref={editorRef}
      document={file}
      mode={{ kind: "edit", author: { id: user.id, name: user.name } }}
    />
  );
}
```

`document` accepts a `File`, `Blob`, `ArrayBuffer`, or `Uint8Array`; `mode` selects read-only, comment-only, or full editing and names the author new comments are attributed to; use the ref to export the edited document as bytes.

## What it does

- Provides browser editing with built-in controls, read-only and comment-only modes, and document locking.
- Fits the page to the available width by default and lets readers choose a fixed zoom level.
- Exports edited documents as DOCX bytes or browser downloads.
- Preserves document structures and package parts across edits, and records who wrote each comment.
- Reads document styles, theme fonts, page layout, and CJK font information for display.

See [Feature support](https://docx-editor.portone.io/docs/features) for the support matrix. Build [custom controls](https://docx-editor.portone.io/docs/custom-controls) or use [programmatic DOCX import and export](https://docx-editor.portone.io/docs/core).

The full [documentation](https://docx-editor.portone.io/docs) lives on the documentation site.

## Contributing

Contributions are welcome; see [CONTRIBUTING.md](./CONTRIBUTING.md) to get started.

## License

Licensed under the [Apache License 2.0](./LICENSE).
