import type { BaseLayoutProps } from "fumadocs-ui/layouts/shared";
import { PortOneSymbol } from "@/components/PortOneSymbol";
import { repositoryUrl } from "./library";

export function baseOptions(): BaseLayoutProps {
  return {
    nav: {
      title: (
        <span className="flex items-center gap-2">
          <PortOneSymbol className="size-[18px] text-[var(--brand-portone)] dark:text-white" />
          docx-editor
        </span>
      ),
      url: "/",
    },
    githubUrl: repositoryUrl,
  };
}
