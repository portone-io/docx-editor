import type { Node } from "fumadocs-core/page-tree";
import type { InferPageType } from "fumadocs-core/source";
import { siteUrl } from "@/lib/library";
import { source } from "@/lib/source";

export type DocsPage = InferPageType<typeof source>;

const markdownRoute = "/llms.mdx/docs";

const absolute = (path: string) => new URL(path, siteUrl).toString();

export const markdownPath = (page: DocsPage) =>
  [markdownRoute, ...page.slugs].join("/");

export const absoluteHtmlUrl = (page: DocsPage) => absolute(page.url);

export const absoluteMarkdownUrl = (page: DocsPage) =>
  absolute(markdownPath(page));

function* walk(nodes: Node[]): Generator<DocsPage> {
  for (const node of nodes) {
    if (node.type === "page") {
      const page = source.getNodePage(node);
      if (page) yield page;
    } else if (node.type === "folder") {
      if (node.index) {
        const page = source.getNodePage(node.index);
        if (page) yield page;
      }
      yield* walk(node.children);
    }
  }
}

export const getPagesInTreeOrder = (): DocsPage[] => [
  ...walk(source.getPageTree().children),
];

export async function getMarkdown(page: DocsPage): Promise<string> {
  const body = await page.data.getText("processed");

  return [
    `# ${page.data.title}`,
    `Source: ${absoluteHtmlUrl(page)}`,
    page.data.description,
    body,
  ]
    .map((section) => section?.trim())
    .filter((section) => section)
    .join("\n\n");
}
