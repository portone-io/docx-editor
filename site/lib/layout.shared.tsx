import type { BaseLayoutProps } from "fumadocs-ui/layouts/shared";
import { libraryName, repositoryUrl } from "./library";

export function baseOptions(): BaseLayoutProps {
  return {
    nav: { title: libraryName },
    githubUrl: repositoryUrl,
  };
}
