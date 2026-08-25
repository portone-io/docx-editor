import type { MetadataRoute } from "next";
import { siteUrl } from "@/lib/library";
import { source } from "@/lib/source";

const absolute = (path: string) => new URL(path, siteUrl).toString();

export default function sitemap(): MetadataRoute.Sitemap {
  return [
    { url: absolute("/") },
    ...source.getPages().map((page) => ({ url: absolute(page.url) })),
  ];
}
