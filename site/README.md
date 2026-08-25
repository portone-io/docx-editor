# Site

`@portone/docx-editor-site` is the documentation and landing site for `@portone/docx-editor`: a Next.js App Router application whose docs pages are rendered by Fumadocs from `content/docs`, and whose landing page mounts the live demo.

Run it from the repository root:

```sh
pnpm dev:site
pnpm build:site
```

## The demo fixture

The landing page needs the same `demo.docx` the editor is developed against, and the repository keeps one copy of it under `__fixtures__/`. `scripts/copy-demo-fixture.mjs` copies it to `public/demo.docx` from `predev` and `prebuild`, and that destination is gitignored, so the binary is never committed twice. The page fetches `/demo.docx` at runtime and wraps the response in a `File`.

## Mounting the demo

The demo arrives through a client-only dynamic import. The editor builds a ProseMirror view against the DOM, so it cannot render on the server.

`next.config.mjs` lists `@portone/docx-editor` and `@portone/docx-editor-demo` in `transpilePackages`: both resolve to TypeScript sources through the workspace link rather than to built output.
