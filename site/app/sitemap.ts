import type { MetadataRoute } from "next";
import { lastModified } from "@/lib/last-modified";
import { siteUrl } from "@/lib/library";
import { source } from "@/lib/source";

const entry = (path: string, sourceFile: string | undefined) => ({
  url: new URL(path, siteUrl).toString(),
  lastModified: lastModified(sourceFile),
});

export default function sitemap(): MetadataRoute.Sitemap {
  return [
    entry("/", "app/page.tsx"),
    ...source.getPages().map((page) => entry(page.url, page.absolutePath)),
  ];
}
