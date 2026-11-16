import { existsSync } from "node:fs";
import { dirname, relative, resolve } from "node:path";
import { parse } from "@babel/parser";

export const RELATIVE_SPECIFIER = /^\.\.?(\/|$)/;

/** TypeScript and Node answer a specifier naming a directory with its index module alone */
const DIRECTORY_SPECIFIER = /(^|\/)\.\.?$|\/$/;

const FILE_SUFFIXES = [".ts", ".tsx"];
const INDEX_SUFFIXES = ["/index.ts", "/index.tsx"];

/** The path an emitted file uses to import another one, or null when nothing answers to it */
export function emittedPath(importer, specifier) {
  const base = resolve(dirname(importer), specifier);
  const suffixes = DIRECTORY_SPECIFIER.test(specifier)
    ? INDEX_SUFFIXES
    : [...FILE_SUFFIXES, ...INDEX_SUFFIXES];
  const source = suffixes
    .map((suffix) => `${base}${suffix}`)
    .find((candidate) => existsSync(candidate));
  if (source === undefined) return null;
  const path = relative(dirname(importer), source).replace(/\.tsx?$/, ".js");
  return path.startsWith(".") ? path : `./${path}`;
}

/** The string literal naming another module, wherever a declaration file can hold one */
function moduleSpecifier(node) {
  switch (node.type) {
    case "ImportDeclaration":
    case "ExportNamedDeclaration":
    case "ExportAllDeclaration":
    case "TSImportType":
      return node.source ?? null;
    case "TSExternalModuleReference":
      return node.expression;
    case "TSModuleDeclaration":
      return node.id.type === "StringLiteral" ? node.id : null;
    default:
      return null;
  }
}

function* nodesOf(node) {
  yield node;
  for (const [key, value] of Object.entries(node)) {
    if (key.endsWith("Comments")) continue;
    for (const child of Array.isArray(value) ? value : [value]) {
      if (typeof child?.type === "string") yield* nodesOf(child);
    }
  }
}

/** A parse failure is the parser's, most likely syntax newer than it reads, and says so */
function parseDeclaration(text) {
  try {
    return parse(text, {
      sourceType: "module",
      plugins: [["typescript", { dts: true }], "decoratorAutoAccessors"],
    });
  } catch (cause) {
    throw new Error(`@babel/parser could not parse it: ${cause.message}`, {
      cause,
    });
  }
}

/**
 * Replaces each relative module specifier in a declaration file's text with the path
 * `emittedPathOf` gives for it, leaving every other character as it was. `emittedPathOf`
 * returns null for a specifier it cannot place, and this throws naming that specifier.
 */
export function rewriteDeclarationSpecifiers(text, emittedPathOf) {
  const replacements = [...nodesOf(parseDeclaration(text).program)]
    .map(moduleSpecifier)
    .filter(
      (literal) => literal !== null && RELATIVE_SPECIFIER.test(literal.value)
    )
    .map((literal) => {
      const path = emittedPathOf(literal.value);
      if (path === null)
        throw new Error(`"${literal.value}" resolves to no source file`);
      return { start: literal.start + 1, end: literal.end - 1, path };
    })
    .sort((a, b) => b.start - a.start);
  return replacements.reduce(
    (rewritten, { start, end, path }) =>
      rewritten.slice(0, start) + path + rewritten.slice(end),
    text
  );
}
