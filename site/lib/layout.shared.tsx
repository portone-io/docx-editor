import type { BaseLayoutProps } from "fumadocs-ui/layouts/shared";
import Image from "next/image";
import { repositoryUrl } from "./library";

export function baseOptions(): BaseLayoutProps {
  return {
    nav: {
      title: (
        <span className="flex items-center gap-2">
          <Image alt="" height={18} src="/icon.png" width={18} />
          docx-editor
        </span>
      ),
      url: "/",
    },
    githubUrl: repositoryUrl,
  };
}
