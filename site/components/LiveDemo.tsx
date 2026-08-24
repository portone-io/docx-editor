"use client";

import dynamic from "next/dynamic";
import { useEffect, useState } from "react";

// The editor builds a ProseMirror view against the DOM, so it cannot render on the server
const DocxEditorDemo = dynamic(
  () =>
    import("@portone-io/docx-editor-demo").then(
      (module) => module.DocxEditorDemo
    ),
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

  if (state.status === "loading") return <p>Opening the demo…</p>;
  if (state.status === "failed") {
    return (
      <p role="alert">
        Could not load {DEMO_FILE_NAME}. {state.detail}
      </p>
    );
  }

  return <DocxEditorDemo document={state.file} />;
}
