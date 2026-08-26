import { notFound } from "next/navigation";
import { absoluteHtmlUrl, getMarkdown } from "@/lib/markdown";
import { source } from "@/lib/source";

export const revalidate = false;

export async function GET(
  _request: Request,
  { params }: RouteContext<"/llms.mdx/docs/[[...slug]]">
) {
  const { slug } = await params;
  const page = source.getPage(slug);
  if (!page) notFound();

  return new Response(await getMarkdown(page), {
    headers: {
      "Content-Type": "text/markdown; charset=utf-8",
      Link: `<${absoluteHtmlUrl(page)}>; rel="canonical"`,
    },
  });
}

export function generateStaticParams() {
  return source.generateParams();
}
