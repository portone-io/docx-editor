import { createMDX } from "fumadocs-mdx/next";

const withMDX = createMDX();

/** @type {import('next').NextConfig} */
const config = {
  reactStrictMode: true,
  // Next writes its own AGENTS.md and CLAUDE.md into this folder otherwise, and the
  // repository keeps that guidance at its root
  agentRules: false,
  // The demo is a workspace link resolving to TypeScript sources, and on main so is the
  // library. On the production branch the library is the published package, whose built
  // JavaScript passes through unchanged.
  transpilePackages: ["@portone/docx-editor-demo", "@portone/docx-editor"],
  // The `.md` suffix is the convention llms.txt specifies for a page's Markdown
  // representation, and a dynamic segment cannot carry it without colliding with
  // the docs page route, so it is rewritten onto the route that renders Markdown
  async rewrites() {
    return [
      { source: "/docs.md", destination: "/llms.mdx/docs" },
      { source: "/docs/:path*.md", destination: "/llms.mdx/docs/:path*" },
    ];
  },
  // The Custom controls section was renamed to Editor API, and the READMEs already
  // published to npm link to the old address
  async redirects() {
    return [
      {
        source: "/docs/custom-controls",
        destination: "/docs/editor-api",
        permanent: true,
      },
      {
        source: "/docs/custom-controls.md",
        destination: "/docs/editor-api.md",
        permanent: true,
      },
      {
        source: "/docs/custom-controls/:path*",
        destination: "/docs/editor-api/:path*",
        permanent: true,
      },
    ];
  },
};

export default withMDX(config);
