import { createMDX } from "fumadocs-mdx/next";

const withMDX = createMDX();

/** @type {import('next').NextConfig} */
const config = {
  reactStrictMode: true,
  // Next writes its own AGENTS.md and CLAUDE.md into this folder otherwise, and the
  // repository keeps that guidance at its root
  agentRules: false,
  // Both workspace packages resolve to TypeScript sources through the workspace link
  transpilePackages: ["@portone/docx-editor", "@portone/docx-editor-demo"],
};

export default withMDX(config);
