import {
  BookOpen,
  FileText,
  Image,
  Languages,
  List,
  type LucideIcon,
  MessageSquare,
  ShieldCheck,
  Table,
  Type,
} from "lucide-react";
import Link from "next/link";
import { InstallCommand } from "@/components/InstallCommand";
import { LiveDemo } from "@/components/LiveDemo";
import { PortOneWordmark } from "@/components/PortOneWordmark";
import { docsRoute, libraryName, repositoryUrl } from "@/lib/library";

const ICON_STROKE_WIDTH = 1.6;

const features: { title: string; description: string; icon: LucideIcon }[] = [
  {
    title: "Lossless round trip",
    description:
      "What you never touch is written back from its original XML, not rebuilt.",
    icon: ShieldCheck,
  },
  {
    title: "Text formatting",
    description:
      "Bold, italic, underline, colors, highlights, fonts, and font sizes.",
    icon: Type,
  },
  {
    title: "Paragraphs and lists",
    description:
      "Styles, alignment, indentation, spacing, and nested bullet or numbered lists.",
    icon: List,
  },
  {
    title: "Tables",
    description: "Merge, split, resize, and paginate long tables across pages.",
    icon: Table,
  },
  {
    title: "Comments",
    description: "Anchored threads with replies, resolution, and reopening.",
    icon: MessageSquare,
  },
  {
    title: "Images and links",
    description: "Insert, paste, resize images; create and edit hyperlinks.",
    icon: Image,
  },
  {
    title: "Notes and page furniture",
    description:
      "View footnotes and endnotes, and preview headers, footers, and page numbers.",
    icon: BookOpen,
  },
  {
    title: "IME composition",
    description: "First-class Korean and CJK text input.",
    icon: Languages,
  },
];

const navLinkClassName =
  "font-medium text-[14px] text-[var(--brand-text-muted)] transition-colors hover:text-[var(--brand-text)]";

function BrandMark() {
  return (
    <FileText
      aria-hidden="true"
      className="flex-none text-[var(--brand-primary)]"
      size={22}
      strokeWidth={ICON_STROKE_WIDTH}
    />
  );
}

export default function LandingPage() {
  return (
    <div className="flex flex-1 flex-col bg-[var(--brand-surface)] text-[var(--brand-text)]">
      <header className="border-[var(--brand-border)] border-b bg-[var(--brand-surface)]">
        <div className="mx-auto flex w-full max-w-[1280px] items-center justify-between gap-4 px-4 py-3 sm:px-6">
          <Link className="flex items-center gap-2.5" href="/">
            <BrandMark />
            <span className="font-bold text-[17px]">docx-editor</span>
          </Link>
          <nav className="flex items-center gap-4 sm:gap-6">
            <Link className={navLinkClassName} href={docsRoute}>
              Docs
            </Link>
            <a
              className={`${navLinkClassName} hidden sm:inline-block`}
              href="#features"
            >
              Features
            </a>
            <a
              className={`${navLinkClassName} hidden sm:inline-block`}
              href={repositoryUrl}
            >
              GitHub
            </a>
            <Link
              className="rounded-[8px] bg-[var(--brand-primary)] px-[18px] py-[9px] font-medium text-[14px] text-white transition-colors hover:bg-[var(--brand-primary-hover)]"
              href={docsRoute}
            >
              Get started
            </Link>
          </nav>
        </div>
      </header>

      <main className="flex flex-1 flex-col">
        <section className="hero-backdrop flex flex-col overflow-hidden border-[var(--brand-border)] border-b px-4 pt-16 sm:px-6 sm:pt-24">
          <div className="mx-auto flex max-w-[720px] flex-col items-center text-center">
            <h1 className="font-extrabold text-[30px] leading-[1.06] tracking-[-0.03em] sm:text-[44px] lg:text-[56px]">
              A simple DOCX editor
              <br />
              for your React app
            </h1>
            <p className="mt-5 max-w-[560px] text-[16px] text-[var(--brand-text-muted)] leading-[1.6]">
              Open, edit, and save Word documents in the browser. Built directly
              on OOXML, so what you don&apos;t edit survives the round trip.
            </p>
            <div className="mt-8">
              <InstallCommand command={`npm i ${libraryName}`} />
            </div>
          </div>
          {/* `.demo` fills its container and brings its own header strip, so the
              card is just the bordered box that carries the height */}
          <div
            className="mx-auto mt-14 h-[540px] w-full max-w-[1020px] overflow-hidden rounded-t-[8px] border border-[var(--brand-border)] border-b-0 bg-[var(--brand-surface)] sm:mt-20 sm:h-[620px] lg:h-[720px]"
            id="demo"
          >
            <LiveDemo />
          </div>
        </section>

        <section
          className="mx-auto w-full max-w-[1280px] scroll-mt-8 px-4 py-16 sm:px-6 sm:py-24"
          id="features"
        >
          <p className="text-center font-mono text-[12px] text-[var(--brand-primary)] tracking-[0.12em]">
            FEATURES
          </p>
          <div className="mt-8 grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
            {features.map(({ icon: FeatureIcon, ...feature }) => (
              <div
                className="rounded-[8px] border border-[var(--brand-border)] bg-[var(--brand-surface)] p-6"
                key={feature.title}
              >
                <FeatureIcon
                  aria-hidden="true"
                  className="text-[var(--brand-primary)]"
                  size={22}
                  strokeWidth={ICON_STROKE_WIDTH}
                />
                <h2 className="mt-4 font-semibold text-[15px]">
                  {feature.title}
                </h2>
                <p className="mt-2 text-[13.5px] text-[var(--brand-text-subtle)] leading-[1.55]">
                  {feature.description}
                </p>
              </div>
            ))}
          </div>
        </section>
      </main>

      <footer className="border-[var(--brand-border)] border-t">
        <div className="mx-auto flex w-full max-w-[1280px] flex-col items-center justify-between gap-6 px-4 py-10 sm:flex-row sm:gap-4 sm:px-6">
          <div className="flex flex-col items-center gap-2 sm:items-start">
            <PortOneWordmark />
            <p className="text-[12px] text-[var(--brand-text-subtle)]">
              Copyright © 2026 PortOne Korea Corp. All rights reserved.
            </p>
          </div>
          <p className="font-mono text-[12px] text-[var(--brand-text-subtle)]">
            Apache-2.0
          </p>
        </div>
      </footer>
    </div>
  );
}
