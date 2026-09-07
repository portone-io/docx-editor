# Site

`@portone/docx-editor-site` is the documentation and landing site for `@portone/docx-editor`: a Next.js App Router application whose docs pages are rendered by Fumadocs from `content/docs`, and whose landing page mounts the live demo.

It is published at [docx-editor.portone.io](https://docx-editor.portone.io).

Run it from the repository root:

```sh
pnpm dev:site
pnpm build:site
```

## The demo fixture

The landing page needs the same `demo.docx` the editor is developed against, and the repository keeps one copy of it under `__fixtures__/`. `scripts/copy-demo-fixture.mjs` copies it to `public/demo.docx` from `predev` and `prebuild`, and that destination is gitignored, so the binary is never committed twice. The page fetches `/demo.docx` at runtime and wraps the response in a `File`.

## Mounting the demo

The demo arrives through a client-only dynamic import. The editor builds a ProseMirror view against the DOM, so it cannot render on the server.

`next.config.mjs` lists `@portone/docx-editor-demo` in `transpilePackages`, because that package is a workspace link resolving to TypeScript sources.
The library it imports is deliberately not listed: the site installs it from npm, and the published package ships built JavaScript.

## The version the demo runs

`site/package.json` and `demo/package.json` pin `@portone/docx-editor` to an exact released version rather than to `workspace:*`, and the `v{version}` badge beside the demo reads that same pin.
A visitor who tries the demo and then installs the version the badge names gets the behavior they just used.

Both packages have to name the same version, and the demo cannot be left on `workspace:*`: the site depends on the demo, so the library would arrive twice in one build, once from npm and once from the working tree.

`pnpm check:demo-library` holds that together, and `pnpm check` runs it.
It fails when the two pins disagree, when either is a range rather than a release, when the lockfile has not caught up with a changed pin, when the dependency resolves to this repository's own package, and when a release has arrived that the demo is not running yet.
`pnpm pin:demo-library` moves both pins to the newest release and installs it.

`pnpm changeset:version` runs that same command, so preparing a release also picks up a pin that the previous release left behind.
It cannot pin the version being prepared, because that version reaches the registry only when the release publishes, and a dependency on a version npm does not have yet leaves the whole workspace uninstallable.
So the pin moves after a publish rather than with it, and `pnpm check:demo-library` is what reports the gap in between.

`pnpm-workspace.yaml` excludes the library from pnpm's release-maturity delay.
That delay is there to let a compromised third-party release be yanked before anything depends on it, and applying it to this repository's own package would hold the demo a day behind every release of the library it demonstrates.

### Working on the library itself

`pnpm dev:site` serves the released library, so an unreleased change under `src/` does not show up in the site's demo.
Library work belongs in `pnpm dev`, which serves the same demo component through Vite: `demo/vite.config.ts` aliases every `@portone/docx-editor` entry point to the sources it is built from, so edits under `src/` reload there.
Docs pages under `content/docs` are site files and still hot-reload under `pnpm dev:site`.

To make the site itself read the working tree for a one-off check, put `workspace:*` back in `site/package.json` and `demo/package.json`, add `@portone/docx-editor` to `transpilePackages`, and run `pnpm install`.
Revert all three before committing; `pnpm check` fails while they are in place.
A pnpm `overrides` entry does not do this, because it does not displace a direct dependency that already resolves.

## Markdown for AI agents

Every docs page is also served as Markdown at `<page url>.md`, so an agent can read the documentation without recovering prose from markup.
`/llms.txt` indexes those Markdown pages in navigation order, and each HTML page points at its own Markdown with `rel="alternate"`, so an agent can find them either way.
`next.config.mjs` rewrites the `.md` suffix onto `/llms.mdx/docs`, the route that renders a page as Markdown, because a dynamic segment cannot carry the suffix without colliding with the docs page route.
The Markdown body comes from Fumadocs' `includeProcessedMarkdown` postprocess option, which `lib/source.ts` enables so a page can return its processed body.
Each Markdown response carries an HTTP `Link` canonical naming its HTML page, because the two URLs serve the same content and only the HTML one should be indexed.
