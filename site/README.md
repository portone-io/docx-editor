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

`next.config.mjs` lists `@portone/docx-editor-demo` and `@portone/docx-editor` in `transpilePackages`.
The demo is a workspace link resolving to TypeScript sources, and on `main` so is the library.
On the `production` branch the library is the published package, whose built JavaScript passes through unchanged.

## The version the demo runs

| Command or deployment | Site and docs | Editor library |
| --- | --- | --- |
| `pnpm dev` / `pnpm build:demo` | Development demo | Current `src/` |
| `pnpm dev:site` / `pnpm build:site` / Vercel preview of `main` | Checked-out site and docs | The `src/` of the same commit |
| Vercel production, from the `production` branch | Site and docs as of the latest release | That release, installed from npm |

The version badge names the library the demo below it runs: the version in the root `package.json` on `main`, the published release on `production`.
On `main` that number is the last release, so a preview can run sources newer than its badge says.
`pnpm check:demo-library` confirms that the site and the demo agree on where the library comes from and that the installation matches.

### Automatic updates after publishing

After npm publishing succeeds, [Update site release](../.github/workflows/site-release.yml) checks out the release commit, installs the published version into the site and the demo, builds the site, and writes the result to the `production` branch as the release commit plus one commit updating the two manifests and lockfile. Vercel's production deployment must track `production`; `main` deploys as a preview. The live site therefore changes only at a release, and a docs-only fix reaches it with the next one. Failed updates leave the previous release in place.

Nothing moves on `main` after a release: its site keeps running the sources.

If the update fails, run **Update site release** on `main` in GitHub Actions with the already published version. It rebuilds `production` from scratch, so rerunning is always safe. If `production` is right but its deployment failed, retry in Vercel. Neither requires republishing npm.

To build the site against a specific release locally, run `pnpm pin:demo-library 0.3.0` followed by `pnpm build:site`. If installation fails, run `pnpm install` before retrying.
To return to the sources, run `git checkout -- site/package.json demo/package.json pnpm-lock.yaml && pnpm install`.

## Markdown for AI agents

Every docs page is also served as Markdown at `<page url>.md`, so an agent can read the documentation without recovering prose from markup.
`/llms.txt` indexes those Markdown pages in navigation order, and each HTML page points at its own Markdown with `rel="alternate"`, so an agent can find them either way.
`next.config.mjs` rewrites the `.md` suffix onto `/llms.mdx/docs`, the route that renders a page as Markdown, because a dynamic segment cannot carry the suffix without colliding with the docs page route.
The Markdown body comes from Fumadocs' `includeProcessedMarkdown` postprocess option, which `lib/source.ts` enables so a page can return its processed body.
Each Markdown response carries an HTTP `Link` canonical naming its HTML page, because the two URLs serve the same content and only the HTML one should be indexed.
