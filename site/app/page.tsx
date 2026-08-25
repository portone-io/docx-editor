import {
  Languages,
  Lock,
  type LucideIcon,
  Scale,
  ShieldCheck,
  Type,
} from "lucide-react";
import NextImage from "next/image";
import Link from "next/link";
import { InstallCommand } from "@/components/InstallCommand";
import { LiveDemo } from "@/components/LiveDemo";
import { PortOneWordmark } from "@/components/PortOneWordmark";
import { docsRoute, libraryName, repositoryUrl } from "@/lib/library";

const ICON_STROKE_WIDTH = 1.6;

const features: { title: string; description: string; icon: LucideIcon }[] = [
  {
    title: "The editing surface",
    description:
      "Format text, build lists and tables, place images and links, and comment in anchored threads - the everyday edits, applied in place.",
    icon: Type,
  },
  {
    title: "Open source, all of it",
    description:
      "Every feature ships in one package under Apache-2.0 - no paid tier.",
    icon: Scale,
  },
  {
    title: "CJK input",
    description:
      "IME composition is tested against a real browser on every change, so input never drops or doubles.",
    icon: Languages,
  },
  {
    title: "Documents stay in the browser",
    description:
      "Files are opened, edited, and saved entirely on the page - nothing is uploaded to a server or handed to a third party.",
    icon: Lock,
  },
  {
    title: "No silent rewrites",
    description:
      "The editor writes the same OOXML it read, so a document is never converted, rebuilt, or quietly reformatted on the way through.",
    icon: ShieldCheck,
  },
];

const navLinkClassName =
  "font-medium text-[14px] text-[var(--brand-text-muted)] transition-colors hover:text-[var(--brand-text)]";

function BrandMark() {
  return (
    <NextImage
      alt=""
      className="flex-none"
      height={20}
      src="/icon.png"
      width={20}
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
              href={repositoryUrl}
              rel="noreferrer"
              target="_blank"
            >
              GitHub
            </a>
            <Link
              className="rounded-[8px] bg-[var(--brand-primary)] px-[18px] py-[9px] font-medium text-[14px] text-white transition-colors hover:bg-[var(--brand-primary-hover)]"
              href={`${docsRoute}/getting-started`}
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
              <span aria-hidden="true" className="hero-caret" />
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
            className="mx-auto mt-14 h-[540px] w-full max-w-[1240px] overflow-hidden rounded-t-[8px] border border-[var(--brand-border)] border-b-0 bg-[var(--brand-surface)] sm:mt-20 sm:h-[620px] lg:h-[720px]"
            id="demo"
          >
            <LiveDemo />
          </div>
        </section>

        <section
          className="mx-auto w-full max-w-[1280px] scroll-mt-8 px-6 py-16 sm:px-12 sm:py-24 lg:px-20"
          id="features"
        >
          <div className="flex flex-col gap-10 lg:flex-row lg:gap-12">
            <div className="lg:w-[340px] lg:flex-none">
              <p className="font-mono text-[12px] text-[var(--brand-primary)] tracking-[0.12em]">
                FEATURES
              </p>
              <h2 className="mt-3 font-bold text-[30px] leading-[1.2] tracking-[-0.02em]">
                Simple on purpose
              </h2>
              <p className="mt-3 text-[14.5px] text-[var(--brand-text-subtle)] leading-[1.6]">
                It doesn&apos;t do everything Word does. The features it has
                work without surprises.
              </p>
            </div>
            <div className="flex flex-1 flex-col">
              {features.map(({ icon: FeatureIcon, ...feature }) => (
                <div
                  className="flex items-start gap-3 border-[var(--brand-border)] border-b py-5"
                  key={feature.title}
                >
                  <FeatureIcon
                    aria-hidden="true"
                    className="mt-0.5 flex-none text-[var(--brand-primary)]"
                    size={20}
                    strokeWidth={ICON_STROKE_WIDTH}
                  />
                  <div>
                    <h3 className="font-semibold text-[15px]">
                      {feature.title}
                    </h3>
                    <p className="mt-1 max-w-[68ch] text-[14px] text-[var(--brand-text-subtle)] leading-[1.6]">
                      {feature.description}
                    </p>
                  </div>
                </div>
              ))}
            </div>
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
