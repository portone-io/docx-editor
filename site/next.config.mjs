import { createMDX } from "fumadocs-mdx/next";

const withMDX = createMDX();

/** @type {import('next').NextConfig} */
const config = {
  reactStrictMode: true,
  // Next writes its own AGENTS.md and CLAUDE.md into this folder otherwise, and the
  // repository keeps that guidance at its root
  agentRules: false,
  // The demo is a workspace package and resolves to TypeScript sources. The library it imports
  // is the published package, which ships built JavaScript and needs no transpiling.
  transpilePackages: ["@portone/docx-editor-demo"],
  // The `.md` suffix is the convention llms.txt specifies for a page's Markdown
  // representation, and a dynamic segment cannot carry it without colliding with
  // the docs page route, so it is rewritten onto the route that renders Markdown
  async rewrites() {
    return [
      { source: "/docs.md", destination: "/llms.mdx/docs" },
      { source: "/docs/:path*.md", destination: "/llms.mdx/docs/:path*" },
    ];
  },
};

export default withMDX(config);
