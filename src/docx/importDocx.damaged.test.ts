// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";
import {
  fixtureNames,
  producerFixtureNames,
  readFixture,
  readProducerFixture,
} from "../__testing__/docx";
import { editorStateForSession } from "../editor/createEditor";
import { DocxExportError, DocxImportError } from "../ooxml/errors";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import { CONTENT_TYPES_PATH, markupParts } from "./packageParts";

const CASE_BUDGET_MS = 10_000;

type Random = () => number;

function fnv1a(text: string): number {
  let hash = 0x811c9dc5;
  for (const byte of new TextEncoder().encode(text)) {
    hash = Math.imul(hash ^ byte, 0x01000193) >>> 0;
  }
  return hash;
}

/** xorshift32 has a fixed point at zero, so a zero seed is moved off it */
function xorshift32(seed: number): Random {
  let state = seed || 0x9e3779b9;
  return () => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    state >>>= 0;
    return state;
  };
}

function truncated(bytes: Uint8Array, share: number): Uint8Array<ArrayBuffer> {
  return bytes.slice(0, Math.floor(bytes.length * share));
}

function flipped(
  bytes: Uint8Array,
  within: number,
  random: Random
): Uint8Array<ArrayBuffer> {
  const damaged = bytes.slice();
  const span = Math.min(within, damaged.length);
  for (let flip = 0; flip < 8; flip += 1) {
    damaged[random() % span] ^= 1 << (random() % 8);
  }
  return damaged;
}

function withPart(
  bytes: Uint8Array,
  path: string,
  damage: (part: Uint8Array) => Uint8Array<ArrayBuffer> | null
): Uint8Array {
  const parts = unzipSync(bytes);
  const damaged = damage(parts[path]);
  if (damaged === null) delete parts[path];
  else parts[path] = damaged;
  return zipSync(parts, { level: 6, mtime: new Date(2026, 0, 1) });
}

interface Variant {
  readonly name: string;
  readonly damage: (bytes: Uint8Array, random: Random) => Uint8Array;
}

function variantsOf(bytes: Uint8Array): Variant[] {
  const { mainPartPath } = importDocx(bytes).session;
  const parts = unzipSync(bytes);
  const otherMarkupParts = Array.from(
    markupParts(Object.keys(parts), parts[CONTENT_TYPES_PATH])
  ).filter((path) => path !== mainPartPath);
  return [
    { name: "truncate-50", damage: (zip) => truncated(zip, 0.5) },
    { name: "truncate-90", damage: (zip) => truncated(zip, 0.9) },
    {
      name: "flip-header",
      damage: (zip, random) => flipped(zip, 512, random),
    },
    {
      name: "flip-body",
      damage: (zip, random) => flipped(zip, zip.length, random),
    },
    {
      name: "main-part-truncated",
      damage: (zip) =>
        withPart(zip, mainPartPath, (part) => truncated(part, 0.5)),
    },
    {
      name: "main-part-flipped",
      damage: (zip, random) =>
        withPart(zip, mainPartPath, (part) =>
          flipped(part, part.length, random)
        ),
    },
    {
      name: "main-part-missing",
      damage: (zip) => withPart(zip, mainPartPath, () => null),
    },
    ...otherMarkupParts.map(
      (path): Variant => ({
        name: `part-truncated:${path}`,
        damage: (zip) => withPart(zip, path, (part) => truncated(part, 0.5)),
      })
    ),
  ];
}

function openedOrRefused(
  bytes: Uint8Array
): ReturnType<typeof importDocx> | null {
  try {
    return importDocx(bytes);
  } catch (error) {
    if (error instanceof DocxImportError) return null;
    throw error;
  }
}

function writtenOrRefused(run: () => Uint8Array): void {
  try {
    run();
  } catch (error) {
    if (!(error instanceof DocxExportError)) throw error;
  }
}

function openAndExport(bytes: Uint8Array): void {
  const opened = openedOrRefused(bytes);
  if (opened === null) return;
  writtenOrRefused(() => exportDocx(opened.doc, opened.session));
  const state = editorStateForSession(opened);
  writtenOrRefused(() => exportDocx(state.doc, opened.session));
}

const fixtures = [
  ...fixtureNames.map((name) => ({ name, bytes: readFixture(name) })),
  ...producerFixtureNames.map((name) => ({
    name: `producers/${name}`,
    bytes: readProducerFixture(name),
  })),
];

describe("opening a damaged package", () => {
  for (const fixture of fixtures) {
    for (const variant of variantsOf(fixture.bytes)) {
      it(
        `${fixture.name} · ${variant.name}`,
        () => {
          const random = xorshift32(fnv1a(`${fixture.name}/${variant.name}`));
          openAndExport(variant.damage(fixture.bytes, random));
        },
        CASE_BUDGET_MS
      );
    }
  }
});
