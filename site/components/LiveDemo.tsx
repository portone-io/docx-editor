"use client";

import dynamic from "next/dynamic";
import { type ReactNode, useEffect, useState } from "react";

// The editor builds a ProseMirror view against the DOM, so it cannot render on the server
const DocxEditorDemo = dynamic(
  () =>
    import("@portone/docx-editor-demo").then((module) => module.DocxEditorDemo),
  { ssr: false }
);

const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const DEMO_FILE_NAME = "demo.docx";

type DemoState =
  | { status: "loading" }
  | { status: "ready"; file: File }
  | { status: "failed"; detail: string };

async function loadDemoFile(): Promise<File> {
  const response = await fetch(`/${DEMO_FILE_NAME}`);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const blob = await response.blob();
  return new File([blob], DEMO_FILE_NAME, { type: DOCX_MIME_TYPE });
}

// Mirrors the demo's own header strip, so mounting the editor does not shift
// the card's contents
function DemoPlaceholder({
  children,
  role,
}: {
  children: ReactNode;
  role?: "alert";
}) {
  return (
    <div className="flex h-full flex-col bg-[var(--brand-surface)]">
      <p className="min-h-[44px] border-[var(--brand-border)] border-b px-3.5 py-2.5 font-mono text-[11px] text-[var(--brand-text-subtle)] uppercase tracking-[0.08em]">
        Live demo
      </p>
      <p
        className="flex flex-1 items-center justify-center px-6 text-center text-[14px] text-[var(--brand-text-subtle)]"
        role={role}
      >
        {children}
      </p>
    </div>
  );
}

export function LiveDemo() {
  const [state, setState] = useState<DemoState>({ status: "loading" });

  useEffect(() => {
    let cancelled = false;
    void loadDemoFile().then(
      (file) => {
        if (!cancelled) setState({ status: "ready", file });
      },
      (error: unknown) => {
        if (cancelled) return;
        setState({
          status: "failed",
          detail: error instanceof Error ? error.message : String(error),
        });
      }
    );
    return () => {
      cancelled = true;
    };
  }, []);

  if (state.status === "loading") {
    return <DemoPlaceholder>Opening the demo…</DemoPlaceholder>;
  }
  if (state.status === "failed") {
    return (
      <DemoPlaceholder role="alert">
        Could not load {DEMO_FILE_NAME}. {state.detail}
      </DemoPlaceholder>
    );
  }

  return <DocxEditorDemo document={state.file} />;
}
