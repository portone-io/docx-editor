# Site

`@portone/docx-editor-site` is the documentation and landing site for `@portone/docx-editor`: a Next.js App Router application whose docs pages are rendered by Fumadocs from `content/docs`, and whose landing page mounts the live demo.

It is published at [docx-editor.portone.io](https://docx-editor.portone.io).

Run it from the repository root:

```sh
pnpm dev:site
pnpm build:site
```

## Metadata and indexing

`lib/library.ts` reads the site origin from the `homepage` field of the root `package.json`, so the domain is written once. `app/layout.tsx` turns it into `metadataBase`, and `app/sitemap.ts` and `app/robots.ts` build absolute URLs from it. The sitemap enumerates docs pages through `source.getPages()`, so a new MDX file appears without further edits.

Canonical URLs are set per page rather than in the root layout, because metadata inherits down the route tree and a canonical declared once at the root would point every page at `/`.

## The demo fixture

The landing page needs the same `demo.docx` the editor is developed against, and the repository keeps one copy of it under `__fixtures__/`. `scripts/copy-demo-fixture.mjs` copies it to `public/demo.docx` from `predev` and `prebuild`, and that destination is gitignored, so the binary is never committed twice. The page fetches `/demo.docx` at runtime and wraps the response in a `File`.

## Mounting the demo

The demo arrives through a client-only dynamic import. The editor builds a ProseMirror view against the DOM, so it cannot render on the server.

`next.config.mjs` lists `@portone/docx-editor` and `@portone/docx-editor-demo` in `transpilePackages`: both resolve to TypeScript sources through the workspace link rather than to built output.
