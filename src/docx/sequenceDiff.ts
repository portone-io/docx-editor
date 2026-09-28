/**
 * The elements two sequences share, as the shortest edit between them, and the block-by-block
 * alignment `compareDocx` builds on it.
 *
 * The search is Myers' in linear space (the middle-snake bisection), so a long sequence rewritten
 * throughout costs time rather than a table the size of both sequences multiplied. The time is
 * bounded too: past `COMPARISON_BUDGET` the search gives up and reports everything between the
 * shared head and tail as removed and added, which is a correct edit if not the shortest one.
 */

export type Match = readonly [original: number, revised: number];

/**
 * How many element comparisons and diagonal probes one search may spend. It is a count rather
 * than a deadline so that the same two inputs always get the same answer. Five million keeps a
 * search under a tenth of a second, and still finds a thousand edits scattered through a text of
 * forty thousand characters.
 */
export const COMPARISON_BUDGET = 5_000_000;

type Same = (original: number, revised: number) => boolean;

interface Budget {
  left: number;
}

type Split = readonly [original: number, revised: number];

/** Where the two paths of the bisection meet; null where the ranges share nothing, and `exhausted` where the budget ran out */
function bisect(
  originalFrom: number,
  n: number,
  revisedFrom: number,
  m: number,
  same: Same,
  budget: Budget
): Split | null | "exhausted" {
  const maxD = Math.ceil((n + m) / 2);
  const offset = maxD;
  const length = 2 * maxD + 2;
  const forward = new Int32Array(length).fill(-1);
  const backward = new Int32Array(length).fill(-1);
  forward[offset + 1] = 0;
  backward[offset + 1] = 0;
  const delta = n - m;
  // With an odd difference the forward path is the one to reach the overlap first
  const front = delta % 2 !== 0;
  let forwardStart = 0;
  let forwardEnd = 0;
  let backwardStart = 0;
  let backwardEnd = 0;
  for (let d = 0; d < maxD; d += 1) {
    for (let k = -d + forwardStart; k <= d - forwardEnd; k += 2) {
      budget.left -= 1;
      const at = offset + k;
      let x =
        k === -d || (k !== d && forward[at - 1] < forward[at + 1])
          ? forward[at + 1]
          : forward[at - 1] + 1;
      let y = x - k;
      const start = x;
      while (x < n && y < m && same(originalFrom + x, revisedFrom + y)) {
        x += 1;
        y += 1;
      }
      budget.left -= x - start;
      forward[at] = x;
      if (x > n) forwardEnd += 2;
      else if (y > m) forwardStart += 2;
      else if (front) {
        const other = offset + delta - k;
        if (other >= 0 && other < length && backward[other] !== -1) {
          if (x >= n - backward[other]) return [x, y];
        }
      }
    }
    for (let k = -d + backwardStart; k <= d - backwardEnd; k += 2) {
      budget.left -= 1;
      const at = offset + k;
      let x =
        k === -d || (k !== d && backward[at - 1] < backward[at + 1])
          ? backward[at + 1]
          : backward[at - 1] + 1;
      let y = x - k;
      const start = x;
      while (
        x < n &&
        y < m &&
        same(originalFrom + n - x - 1, revisedFrom + m - y - 1)
      ) {
        x += 1;
        y += 1;
      }
      budget.left -= x - start;
      backward[at] = x;
      if (x > n) backwardEnd += 2;
      else if (y > m) backwardStart += 2;
      else if (!front) {
        const other = offset + delta - k;
        if (other >= 0 && other < length && forward[other] !== -1) {
          const meet = forward[other];
          if (meet >= n - x) return [meet, offset + meet - other];
        }
      }
    }
    if (budget.left <= 0) return "exhausted";
  }
  return null;
}

/** Appends the matches of the two ranges to `found` in order, and answers false once the budget runs out */
function matchesIn(
  originalFrom: number,
  originalTo: number,
  revisedFrom: number,
  revisedTo: number,
  same: Same,
  budget: Budget,
  found: Match[]
): boolean {
  while (
    originalFrom < originalTo &&
    revisedFrom < revisedTo &&
    same(originalFrom, revisedFrom)
  ) {
    found.push([originalFrom, revisedFrom]);
    originalFrom += 1;
    revisedFrom += 1;
  }
  const tail: Match[] = [];
  while (
    originalFrom < originalTo &&
    revisedFrom < revisedTo &&
    same(originalTo - 1, revisedTo - 1)
  ) {
    originalTo -= 1;
    revisedTo -= 1;
    tail.push([originalTo, revisedTo]);
  }
  if (originalFrom < originalTo && revisedFrom < revisedTo) {
    const split = bisect(
      originalFrom,
      originalTo - originalFrom,
      revisedFrom,
      revisedTo - revisedFrom,
      same,
      budget
    );
    if (split === "exhausted") return false;
    if (split !== null) {
      const [x, y] = split;
      const kept =
        matchesIn(
          originalFrom,
          originalFrom + x,
          revisedFrom,
          revisedFrom + y,
          same,
          budget,
          found
        ) &&
        matchesIn(
          originalFrom + x,
          originalTo,
          revisedFrom + y,
          revisedTo,
          same,
          budget,
          found
        );
      if (!kept) return false;
    }
  }
  found.push(...tail.reverse());
  return true;
}

/**
 * The index pairs of a longest common subsequence of two sequences under `same`, in order.
 *
 * Where the search runs past its budget, only the shared head and tail are matched. Searches that
 * pass one budget between them share it, so a caller running many bounds them all together.
 */
export function commonSubsequence(
  originalLength: number,
  revisedLength: number,
  same: Same,
  budget: Budget = { left: COMPARISON_BUDGET }
): Match[] {
  let head = 0;
  while (head < originalLength && head < revisedLength && same(head, head)) {
    head += 1;
  }
  let tail = 0;
  while (
    tail < originalLength - head &&
    tail < revisedLength - head &&
    same(originalLength - 1 - tail, revisedLength - 1 - tail)
  ) {
    tail += 1;
  }
  const matched = (from: number, to: number, shift: number): Match[] =>
    Array.from({ length: to - from }, (_, at) => [
      from + at,
      from + at + shift,
    ]);
  const middle: Match[] = [];
  const found = matchesIn(
    head,
    originalLength - tail,
    head,
    revisedLength - tail,
    same,
    budget,
    middle
  );
  return [
    ...matched(0, head, 0),
    ...(found ? middle : []),
    ...matched(
      originalLength - tail,
      originalLength,
      revisedLength - originalLength
    ),
  ];
}

/** One element of a sequence to align: what it would be written as, and what it reads as */
export interface Alignable {
  identity: string;
  text: string;
}

export type Alignment =
  | { kind: "kept"; original: number; revised: number }
  | {
      kind: "changed";
      original: number;
      revised: number;
      change: "text" | "formatting";
    }
  | { kind: "removed"; original: number }
  | { kind: "added"; revised: number };

/** Each distinct key as a small number, so the search compares numbers rather than long strings */
function keyed(
  original: readonly string[],
  revised: readonly string[]
): [Int32Array, Int32Array] {
  const ids = new Map<string, number>();
  const idOf = (key: string) => {
    const known = ids.get(key);
    if (known !== undefined) return known;
    ids.set(key, ids.size);
    return ids.size - 1;
  };
  return [Int32Array.from(original, idOf), Int32Array.from(revised, idOf)];
}

/** The pairs `key` matches in each stretch between two anchors, the anchors kept, in order */
function matchBetween<T>(
  anchors: readonly Match[],
  original: readonly T[],
  revised: readonly T[],
  key: (item: T) => string
): Match[] {
  const [was, now] = keyed(original.map(key), revised.map(key));
  const budget: Budget = { left: COMPARISON_BUDGET };
  const found: Match[] = [];
  let from = 0;
  let to = 0;
  for (const anchor of [
    ...anchors,
    [original.length, revised.length] as const,
  ]) {
    const matches = commonSubsequence(
      anchor[0] - from,
      anchor[1] - to,
      (i, j) => was[from + i] === now[to + j],
      budget
    );
    for (const [i, j] of matches) found.push([from + i, to + j]);
    if (anchor[0] < original.length) found.push(anchor);
    from = anchor[0] + 1;
    to = anchor[1] + 1;
  }
  return found;
}

/**
 * The elements left between two anchors, paired in order: each one on the revised side takes the
 * first element of its kind on the original side after the last one taken, and the rest are
 * removed and added where they stand.
 */
function pairLeftovers<T>(
  original: readonly T[],
  revised: readonly T[],
  kindOf: (item: T) => string,
  [originalFrom, originalTo]: Match,
  [revisedFrom, revisedTo]: Match
): Alignment[] {
  const steps: Alignment[] = [];
  let cursor = originalFrom;
  let added = revisedFrom;
  for (let at = revisedFrom; at < revisedTo; at += 1) {
    const kind = kindOf(revised[at]);
    let partner = cursor;
    while (partner < originalTo && kindOf(original[partner]) !== kind) {
      partner += 1;
    }
    if (partner === originalTo) continue;
    for (; cursor < partner; cursor += 1) {
      steps.push({ kind: "removed", original: cursor });
    }
    for (; added < at; added += 1)
      steps.push({ kind: "added", revised: added });
    steps.push({
      kind: "changed",
      original: partner,
      revised: at,
      change: "text",
    });
    cursor = partner + 1;
    added = at + 1;
  }
  for (; cursor < originalTo; cursor += 1) {
    steps.push({ kind: "removed", original: cursor });
  }
  for (; added < revisedTo; added += 1)
    steps.push({ kind: "added", revised: added });
  return steps;
}

/**
 * Lines two sequences up: elements written the same are kept, then elements reading the same are
 * changed in formatting, and the elements left between two of those are paired as changed in
 * text (`pairLeftovers`). Elements of different kinds are never paired.
 */
export function alignSequences<T extends Alignable>(
  original: readonly T[],
  revised: readonly T[],
  kindOf: (item: T) => string = () => ""
): Alignment[] {
  const exact = matchBetween(
    [],
    original,
    revised,
    (item) => `${kindOf(item)}\u0000${item.identity}`
  );
  const matched = matchBetween(
    exact,
    original,
    revised,
    (item) => `${kindOf(item)}\u0000${item.text}`
  );
  const alignment: Alignment[] = [];
  let from = 0;
  let to = 0;
  for (const [was, now] of [
    ...matched,
    [original.length, revised.length] as const,
  ]) {
    alignment.push(
      ...pairLeftovers(original, revised, kindOf, [from, was], [to, now])
    );
    if (was < original.length) {
      alignment.push(
        original[was].identity === revised[now].identity
          ? { kind: "kept", original: was, revised: now }
          : {
              kind: "changed",
              original: was,
              revised: now,
              change: "formatting",
            }
      );
    }
    from = was + 1;
    to = now + 1;
  }
  return alignment;
}
