import { libraryDescription, repositoryUrl } from "@/lib/library";
import { absoluteMarkdownUrl, getPagesInTreeOrder } from "@/lib/markdown";

export const revalidate = false;

export function GET() {
  const entries = getPagesInTreeOrder().map((page) => {
    const link = `- [${page.data.title}](${absoluteMarkdownUrl(page)})`;
    return page.data.description ? `${link}: ${page.data.description}` : link;
  });

  const body = [
    "# docx-editor",
    `> ${libraryDescription}`,
    "`@portone/docx-editor` opens a Word document in the browser, lets a person edit it, and writes the same OOXML back. Every link below serves a page as Markdown; drop the `.md` suffix for the HTML version.",
    "## Docs",
    entries.join("\n"),
    "## Optional",
    `- [Source code](${repositoryUrl}): the repository, issues, and changelog.`,
  ].join("\n\n");

  return new Response(`${body}\n`, {
    headers: { "Content-Type": "text/plain; charset=utf-8" },
  });
}
