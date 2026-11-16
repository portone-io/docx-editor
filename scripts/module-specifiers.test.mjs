import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { test } from "node:test";
import {
  emittedPath,
  rewriteDeclarationSpecifiers,
} from "./module-specifiers.mjs";

const directories = new Set([".", "..", "./table", "../commands"]);

/** Places a specifier the way the build does for a module beside it, recording each request */
function resolver() {
  const asked = [];
  const emittedPathOf = (specifier) => {
    asked.push(specifier);
    if (specifier.includes("missing")) return null;
    if (directories.has(specifier))
      return specifier === "." ? "./index.js" : `${specifier}/index.js`;
    return `${specifier}.js`;
  };
  return { asked, emittedPathOf };
}

const rewrite = (text) => {
  const { asked, emittedPathOf } = resolver();
  return { text: rewriteDeclarationSpecifiers(text, emittedPathOf), asked };
};

test("each way a declaration names a module beside it gets the emitted path", () => {
  const cases = [
    ['import { a } from "./a";', 'import { a } from "./a.js";'],
    ['import type { B } from "./b";', 'import type { B } from "./b.js";'],
    ['import "./side";', 'import "./side.js";'],
    ['export { c } from "./c";', 'export { c } from "./c.js";'],
    ['export type { D } from "./d";', 'export type { D } from "./d.js";'],
    ['export * from "./e";', 'export * from "./e.js";'],
    ['export * as f from "./f";', 'export * as f from "./f.js";'],
    ['type G = import("./g").G;', 'type G = import("./g.js").G;'],
    ['type H = typeof import("./h");', 'type H = typeof import("./h.js");'],
    ['import i = require("./i");', 'import i = require("./i.js");'],
    [
      'export import j = require("./j");',
      'export import j = require("./j.js");',
    ],
    ['declare module "./k" {}', 'declare module "./k.js" {}'],
    [
      'declare module "./l" { import { m } from "./m"; }',
      'declare module "./l.js" { import { m } from "./m.js"; }',
    ],
  ];
  for (const [before, after] of cases)
    assert.equal(rewrite(before).text, after);
});

test("each relative specifier in one text becomes the path the resolver gives for it", () => {
  assert.equal(
    rewrite('export * from "./table";\nexport { x } from "../commands";').text,
    'export * from "./table/index.js";\nexport { x } from "../commands/index.js";'
  );
});

test("a parent or own directory named alone is relative and goes to the resolver", () => {
  assert.equal(
    rewrite('type A = import("..").A;\ntype B = import(".").B;').text,
    'type A = import("../index.js").A;\ntype B = import("./index.js").B;'
  );
});

test("a bare specifier is left alone and never handed to the resolver", () => {
  const text = [
    'import { Node } from "prosemirror-model";',
    'import type { ReactNode } from "react";',
    'export * from "react/jsx-runtime";',
    'type S = import("prosemirror-state").EditorState;',
    'declare module "react" { interface X {} }',
    'import r = require(".hidden");',
  ].join("\n");
  const { text: rewritten, asked } = rewrite(text);
  assert.equal(rewritten, text);
  assert.deepEqual(asked, []);
});

test("module-like text in a comment or a string literal type is left alone", () => {
  const text = [
    '// import { a } from "./a";',
    '/** export * from "./b"; */',
    'export type Path = "./c";',
    'export declare const d: "../d";',
  ].join("\n");
  const { text: rewritten, asked } = rewrite(text);
  assert.equal(rewritten, text);
  assert.deepEqual(asked, []);
});

test("text with characters outside ASCII before a specifier keeps its place", () => {
  assert.equal(
    rewrite(
      '/** 계약서 · 문서 👍 */\nexport type T = "한국어";\nexport { a } from "./a";'
    ).text,
    '/** 계약서 · 문서 👍 */\nexport type T = "한국어";\nexport { a } from "./a.js";'
  );
});

test("a relative specifier the resolver cannot place throws naming it", () => {
  assert.throws(
    () => rewrite('export { a } from "./a";\nexport * from "./missing";'),
    /"\.\/missing" resolves to no source file/
  );
});

test("text the parser cannot read throws naming the parser, not a specifier", () => {
  assert.throws(
    () => rewrite('export { a } from "./a";\nexport { from'),
    /^Error: @babel\/parser could not parse it: /
  );
});

test("an accessor field in a declared class does not stop the rewrite", () => {
  assert.equal(
    rewrite(
      'declare class C {\n  accessor count: number;\n}\nexport { a } from "./a";'
    ).text,
    'declare class C {\n  accessor count: number;\n}\nexport { a } from "./a.js";'
  );
});

test("emittedPath names each source module by the file it is emitted as", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "emitted-path-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  for (const file of [
    "a.ts",
    "b.tsx",
    "dir/index.ts",
    "pair.ts",
    "pair/index.ts",
  ]) {
    await mkdir(dirname(join(root, file)), { recursive: true });
    await writeFile(join(root, file), "");
  }
  const cases = [
    ["user.ts", "./a", "./a.js"],
    ["user.ts", "./b", "./b.js"],
    ["user.ts", "./dir", "./dir/index.js"],
    ["user.ts", "./pair", "./pair.js"],
    ["user.ts", "./pair/", "./pair/index.js"],
    ["pair/user.ts", ".", "./index.js"],
    ["pair/deep/user.ts", "..", "../index.js"],
    ["pair/deep/more/user.ts", "../..", "../../index.js"],
    ["pair/deep/user.ts", "./..", "../index.js"],
    ["dir/user.ts", "../pair/.", "../pair/index.js"],
    ["pair/deep/user.ts", "../../a", "../../a.js"],
    ["user.ts", "./missing", null],
    ["user.ts", "./a/", null],
  ];
  for (const [importer, specifier, expected] of cases)
    assert.equal(
      emittedPath(join(root, importer), specifier),
      expected,
      `${specifier} from ${importer}`
    );
});
