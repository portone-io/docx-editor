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

The site demonstrates a published, stable npm release. Its source files and documentation still come from the working tree.

| Command or deployment | Site and docs | Editor library |
| --- | --- | --- |
| `pnpm dev` / `pnpm build:demo` | The development demo shell | Current `src/`, through Vite aliases |
| `pnpm dev:site` | Current working tree, with hot reload | npm `latest`, resolved once at startup |
| `pnpm build:site` / Vercel | The checked-out site and docs | The exact committed pin |

`site/package.json` and `demo/package.json` name the same exact library version, and the badge reads that pin. Both are consumers of the published package; leaving one on `workspace:*` would mix released and unreleased code. The library's JavaScript and CSS resolve from the same installation.

Starting the local site runs `pnpm pin:demo-library` before Next starts. If a newer release exists, it updates both manifests, the lockfile, and the installed package. Those three tracked files may therefore change when starting the site from an older checkout. It does not change `src/` or move the site's documentation to a release tag. The server keeps that version until restarted.

`pnpm check:demo-library` checks agreement between both pins and installed packages and rejects a workspace link. It is offline: a new npm release cannot make an unchanged library commit fail `pnpm check`. Fresh CI installs use the committed lockfile. The root package version may legitimately be ahead of the demo while a release is being prepared.

### Automatic updates after publishing

The [release workflow](../.github/workflows/release.yml) waits for Changesets to finish. When it successfully publishes this package, it passes that exact version to [Update site release](../.github/workflows/site-release.yml). Preparing a release pull request or failing to publish does not start a site update.

The site workflow:

1. Checks out current `main` so it includes the latest site and docs.
2. Waits for the requested version to be readable from npm and installs it for both consumers.
3. Builds the site against that version.
4. Creates one commit containing only the two manifests and lockfile. The existing Vercel Git integration deploys that commit.

Registry reads and installation each have at most six attempts, with five seconds between attempts. Each registry command times out after 15 seconds and each installation after two minutes. The version is resolved once; retries never switch to a newer `latest`. If preparation fails, the workflow creates no commit and the current site's demo version remains in place.

The commit uses GitHub's `expectedHeadOid` check. If `main` advances during the build, it fails rather than overwriting the newer commit. A delayed or manually retried older release cannot downgrade a newer demo pin.

The existing GitHub token creates the version commit; no additional Vercel credential is needed. Vercel's Git integration must remain enabled for production `main`. GitHub does not start another CI or release workflow for this token's commit, so the site build runs before the commit is created. This also avoids a release loop. Vercel performs its own deployment build and retains the previous production deployment if that fails.

### Retrying and checking a specific release

If the site update fails before the commit, run **Update site release** from GitHub Actions on `main`, supplying the already published version. This also recovers a package that reached npm before a later release step failed. It does not republish the package. If the version commit exists and Vercel failed, retry that deployment in Vercel; rerunning the update workflow with unchanged pins does not create an empty commit.

To prepare a specific published version locally:

```sh
pnpm pin:demo-library 0.3.0
pnpm build:site
```

The update command explicitly permits lockfile changes in CI; ordinary installs remain frozen. On failure it restores manifests and lockfile. Run `pnpm install` before retrying to reconcile any partially changed installation. Site startup requires the registry to be available and fails clearly if it cannot resolve a release; it does not fall back to unreleased source code.

`pnpm-workspace.yaml` excludes this repository's own package from pnpm's release-age delay, so the newly published version can be installed immediately.

## Markdown for AI agents

Every docs page is also served as Markdown at `<page url>.md`, so an agent can read the documentation without recovering prose from markup.
`/llms.txt` indexes those Markdown pages in navigation order, and each HTML page points at its own Markdown with `rel="alternate"`, so an agent can find them either way.
`next.config.mjs` rewrites the `.md` suffix onto `/llms.mdx/docs`, the route that renders a page as Markdown, because a dynamic segment cannot carry the suffix without colliding with the docs page route.
The Markdown body comes from Fumadocs' `includeProcessedMarkdown` postprocess option, which `lib/source.ts` enables so a page can return its processed body.
Each Markdown response carries an HTTP `Link` canonical naming its HTML page, because the two URLs serve the same content and only the HTML one should be indexed.
