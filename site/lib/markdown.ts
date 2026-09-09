import type { Node } from "fumadocs-core/page-tree";
import type { InferPageType } from "fumadocs-core/source";
import { siteUrl } from "@/lib/library";
import { source } from "@/lib/source";

export type DocsPage = InferPageType<typeof source>;

const absolute = (path: string) => new URL(path, siteUrl).toString();

export const markdownPath = (page: DocsPage) => `${page.url}.md`;

export const absoluteHtmlUrl = (page: DocsPage) => absolute(page.url);

export const absoluteMarkdownUrl = (page: DocsPage) =>
  absolute(markdownPath(page));

const CODE_FENCE = /^[ \t]*```/;
const CARDS_BLOCK = /^[ \t]*<Cards>[\s\S]*?^[ \t]*<\/Cards>[ \t]*$/gm;
const CARD = /<Card\b(?<attributes>[\s\S]*?)\/>/g;
const CARD_HREF = /\bhref="(?<value>[^"]*)"/;
const CARD_TITLE = /\btitle="(?<value>[^"]*)"/;
const CARD_DESCRIPTION = /\bdescription="(?<value>[^"]*)"/;
const STEPS_BLOCK = /^[ \t]*<Steps>[\s\S]*?^[ \t]*<\/Steps>[ \t]*$/gm;
const STEP_WRAPPER = /^[ \t]*<\/?Steps?>[ \t]*$/;
const DOCS_LINK = /\]\((?<target>\/docs[^\s)]*)\)/g;
// The route path and whatever anchor or query follows it, so the `.md` suffix
// lands on the path rather than at the end of the link.
const DOCS_TARGET = /^(?<path>\/docs(?:\/[^#?]*)?)(?<suffix>[#?].*)?$/;

const markdownTarget = (target: string): string => {
  const groups = target.match(DOCS_TARGET)?.groups;
  return groups?.path
    ? absolute(`${groups.path}.md${groups.suffix ?? ""}`)
    : target;
};

const attributeValue = (attributes: string, pattern: RegExp) =>
  attributes.match(pattern)?.groups?.value;

const cardAsListItem = (attributes: string): string => {
  const href = attributeValue(attributes, CARD_HREF);
  const title = attributeValue(attributes, CARD_TITLE) ?? href ?? "";
  const description = attributeValue(attributes, CARD_DESCRIPTION);
  const link = href ? `[${title}](${markdownTarget(href)})` : title;
  return description ? `- ${link}: ${description}` : `- ${link}`;
};

const cardsAsList = (markdown: string): string =>
  markdown.replace(CARDS_BLOCK, (block) =>
    [...block.matchAll(CARD)]
      .map((card) => cardAsListItem(card.groups?.attributes ?? ""))
      .join("\n")
  );

const dedented = (lines: readonly string[]): string[] => {
  const indents = lines
    .filter((line) => line.trim())
    .map((line) => line.length - line.trimStart().length);
  const width = indents.length > 0 ? Math.min(...indents) : 0;
  return lines.map((line) => line.slice(width));
};

// A step's own content is indented under its wrapper, which Markdown would read
// as an indented code block once the wrapper is gone.
const withoutStepWrappers = (markdown: string): string =>
  markdown.replace(STEPS_BLOCK, (block) => {
    const lines = block.split("\n").filter((line) => !STEP_WRAPPER.test(line));
    return dedented(lines).join("\n");
  });

const absoluteDocLinks = (markdown: string): string =>
  markdown.replace(
    DOCS_LINK,
    (_link, target: string) => `](${markdownTarget(target)})`
  );

const withProse = (
  markdown: string,
  transform: (prose: string) => string
): string => {
  const parts: string[] = [];
  let prose: string[] = [];
  let fenced = false;

  const takeProse = () => {
    if (prose.length > 0) parts.push(transform(prose.join("\n")));
    prose = [];
  };

  for (const line of markdown.split("\n")) {
    if (CODE_FENCE.test(line)) {
      takeProse();
      parts.push(line);
      fenced = !fenced;
    } else if (fenced) {
      takeProse();
      parts.push(line);
    } else {
      prose.push(line);
    }
  }
  takeProse();

  return parts.join("\n");
};

// The step wrappers go first, because a step's fenced samples share its
// indentation and are dedented with the prose around them.
const asPlainMarkdown = (body: string): string =>
  withProse(withoutStepWrappers(body), (prose) =>
    absoluteDocLinks(cardsAsList(prose))
  );

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
  const body = asPlainMarkdown(await page.data.getText("processed"));

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
