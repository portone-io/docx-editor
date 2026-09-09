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

| Command or deployment | Site and docs | Editor library |
| --- | --- | --- |
| `pnpm dev` / `pnpm build:demo` | Development demo | Current `src/` |
| `pnpm dev:site` | Current working tree, with hot reload | npm `latest`, resolved at startup |
| `pnpm build:site` / Vercel preview of `main` | Checked-out site and docs | Exact committed release pin, which a pull request moves to each release |
| Vercel production, from the `production` branch | Site and docs as of the latest release | That release |

The site and demo pin the same published library version, including its CSS and version badge. Local site startup needs network access and may update `site/package.json`, `demo/package.json`, and `pnpm-lock.yaml` when a newer release exists. The running server keeps that version until restarted.

### Automatic updates after publishing

After npm publishing succeeds, [Update site release](../.github/workflows/site-release.yml) checks out the release commit, installs that exact version, builds the site, and writes the result to the `production` branch as the release commit plus one commit updating the two manifests and lockfile. Vercel's production deployment must track `production`; `main` deploys as a preview. The live site therefore changes only at a release, and a docs-only fix reaches it with the next one. Failed updates leave the previous release in place.

The workflow then opens a pull request from `production` to `main`, because nothing else moves the pins there and `main` would otherwise keep building an older library.
`production` is the release commit plus that pin commit, so it is already the head of the pull request, and one left open follows the next release too.
The workflow opens nothing when `main` already pins the released version.
A workflow token opens it, so its checks wait for **Approve workflows to run** in the merge box, the one approval a release still asks of a person.

If the update fails, run **Update site release** on `main` in GitHub Actions with the already published version. It rebuilds `production` from scratch, so rerunning is always safe. If `production` is right but its deployment failed, retry in Vercel. Neither requires republishing npm.

To check a specific release locally, run `pnpm pin:demo-library 0.3.0` followed by `pnpm build:site`. If installation fails, run `pnpm install` before retrying.

## Markdown for AI agents

Every docs page is also served as Markdown at `<page url>.md`, so an agent can read the documentation without recovering prose from markup.
`/llms.txt` indexes those Markdown pages in navigation order, and each HTML page points at its own Markdown with `rel="alternate"`, so an agent can find them either way.
`next.config.mjs` rewrites the `.md` suffix onto `/llms.mdx/docs`, the route that renders a page as Markdown, because a dynamic segment cannot carry the suffix without colliding with the docs page route.
The Markdown body comes from Fumadocs' `includeProcessedMarkdown` postprocess option, which `lib/source.ts` enables so a page can return its processed body.
Each Markdown response carries an HTTP `Link` canonical naming its HTML page, because the two URLs serve the same content and only the HTML one should be indexed.
