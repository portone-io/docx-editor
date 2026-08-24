import Link from "next/link";
import { LiveDemo } from "@/components/LiveDemo";
import { docsRoute, libraryDescription, libraryName } from "@/lib/library";

export default function LandingPage() {
  return (
    <main className="mx-auto flex w-full max-w-5xl flex-1 flex-col gap-8 px-4 py-16">
      <header className="flex flex-col gap-4">
        <h1 className="text-3xl font-bold">{libraryName}</h1>
        <p className="text-fd-muted-foreground">{libraryDescription}</p>
        <Link href={docsRoute} className="font-medium underline">
          Documentation
        </Link>
      </header>
      {/* `.demo` fills its container, so the mount point carries the height */}
      <section
        id="live-demo"
        className="h-[80vh] overflow-hidden rounded-lg border"
      >
        <LiveDemo />
      </section>
    </main>
  );
}
