"use client";

import { Check, Copy } from "lucide-react";
import { useEffect, useState } from "react";

const COPIED_FEEDBACK_MS = 1600;

export function InstallCommand({ command }: { command: string }) {
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    if (!copied) return;
    const timer = setTimeout(() => setCopied(false), COPIED_FEEDBACK_MS);
    return () => clearTimeout(timer);
  }, [copied]);

  const copy = () => {
    void navigator.clipboard.writeText(command).then(
      () => setCopied(true),
      () => setCopied(false)
    );
  };

  return (
    <button
      aria-label={copied ? "Copied" : `Copy: ${command}`}
      className="group flex items-center gap-3 rounded-[8px] bg-[var(--brand-slate-900)] px-4 py-2 font-mono text-[13px] text-[var(--brand-slate-50)] transition-colors hover:bg-[var(--brand-slate-800)] sm:text-[14px]"
      onClick={copy}
      type="button"
    >
      <span>
        <span className="text-[var(--brand-slate-400)]">$ </span>
        {command}
      </span>
      {copied ? (
        <Check
          aria-hidden="true"
          className="text-[var(--brand-purple-40)]"
          size={16}
          strokeWidth={2}
        />
      ) : (
        <Copy
          aria-hidden="true"
          className="text-[var(--brand-slate-400)] transition-colors group-hover:text-[var(--brand-slate-50)]"
          size={16}
          strokeWidth={2}
        />
      )}
      <span aria-live="polite" className="sr-only">
        {copied ? "Copied" : ""}
      </span>
    </button>
  );
}
