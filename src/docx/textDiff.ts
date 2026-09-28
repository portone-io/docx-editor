/**
 * The characters two texts share and the ones each has alone, as the shortest edit between them
 * (`./sequenceDiff`).
 *
 * The texts are compared a grapheme cluster at a time, so an edit never splits a character a
 * reader sees as one: a surrogate pair, a flag, an emoji with its skin tone, a letter with its
 * combining accent.
 */

import { commonSubsequence } from "./sequenceDiff";

/** One stretch of the edit: `kept` and `removed` spell the original, `kept` and `added` the revision */
export type TextEdit =
  | { kind: "kept"; text: string }
  | { kind: "added"; text: string }
  | { kind: "removed"; text: string };

const graphemes = new Intl.Segmenter(undefined, { granularity: "grapheme" });

function charactersOf(text: string): string[] {
  return Array.from(graphemes.segment(text), ({ segment }) => segment);
}

/**
 * The shortest edit from one text to the other. Between two kept stretches the removed text comes
 * first and the added text after it, each as one stretch, and no stretch is empty.
 */
export function diffText(original: string, revised: string): TextEdit[] {
  const was = charactersOf(original);
  const now = charactersOf(revised);
  const edits: TextEdit[] = [];
  const push = (kind: TextEdit["kind"], text: string) => {
    if (text === "") return;
    const last = edits.at(-1);
    if (last?.kind === kind) {
      edits[edits.length - 1] = { kind, text: last.text + text };
    } else {
      edits.push({ kind, text });
    }
  };
  let from = 0;
  let to = 0;
  const matches = commonSubsequence(
    was.length,
    now.length,
    (i, j) => was[i] === now[j]
  );
  for (const [i, j] of [...matches, [was.length, now.length] as const]) {
    push("removed", was.slice(from, i).join(""));
    push("added", now.slice(to, j).join(""));
    if (i < was.length) push("kept", was[i]);
    from = i + 1;
    to = j + 1;
  }
  return edits;
}
