/**
 * Builds a section index over the ECMA-376 prose PDFs, and looks sections up in it.
 *
 * Part 1 alone is about 5,000 pages, so finding the section a code comment cites means either
 * knowing its page already or reading the table of contents every time. The index maps a section
 * number to the PDF page its body starts on, which is the page number `Read` takes.
 *
 * The index is derived from the PDFs and lands beside them under `spec/pdf/`, which is gitignored:
 * the prose is Ecma's to distribute, and a full list of its headings is close enough to the prose
 * that it stays out of the repository too. What is committed is this script and the notes written
 * from what it finds (`spec/notes/`).
 *
 * Usage:
 *   node scripts/index-ooxml-spec.mjs                 build the index
 *   node scripts/index-ooxml-spec.mjs 17.5.2.23       the section with that number
 *   node scripts/index-ooxml-spec.mjs lock            every section whose title holds "lock"
 */

import { execFile } from "node:child_process";
import { readFile, stat, writeFile } from "node:fs/promises";
import { basename, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const run = promisify(execFile);
const packageRoot = fileURLToPath(new URL("..", import.meta.url));
const pdfDir = join(packageRoot, "spec", "pdf");
const indexFile = join(pdfDir, "section-index.tsv");

/** The PDFs `fetch-ooxml-spec.sh` unpacks, and the short name each is indexed under */
const PARTS = [
  {
    part: "part1",
    file: "Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf",
  },
  {
    part: "part4",
    file: "Ecma Office Open XML Part 4 - Transitional Migration Features.pdf",
  },
];

const SECTION_NUMBER = /^(\d+(?:\.\d+)+)$/;
const SECTION_HEADING = /^(\d+(?:\.\d+)+)\s+(\S.*)$/;
/** The dotted leader a table of contents line carries, which is how a TOC page is told from a body page */
const LEADER = "....";

async function missing(path) {
  try {
    await stat(path);
    return false;
  } catch {
    return true;
  }
}

/** The text of one PDF, page by page. pdftotext writes a form feed between pages */
async function pageTexts(file) {
  const { stdout } = await run("pdftotext", ["-q", join(pdfDir, file), "-"], {
    maxBuffer: 1 << 28,
  });
  return stdout.split("\f");
}

/**
 * The title standing under a section number that sits alone on its line.
 * Null where the next line is another number or a table of contents entry, neither of which is a
 * heading this index should hold.
 */
function titleBelow(lines, at) {
  for (const line of lines.slice(at + 1, at + 4)) {
    const text = line.trim();
    if (!text) continue;
    if (SECTION_NUMBER.test(text) || text.includes(LEADER)) return null;
    return text;
  }
  return null;
}

function isTableOfContents(page) {
  let leaders = 0;
  for (const line of page.split("\n")) {
    if (line.includes(LEADER)) leaders += 1;
    if (leaders >= 3) return true;
  }
  return false;
}

/** Every section heading in this part, as `{ section, page, title }`, the first occurrence winning */
function headingsOf(pages) {
  const found = new Map();
  pages.forEach((page, index) => {
    if (isTableOfContents(page)) return;
    const lines = page.split("\n");
    lines.forEach((line, at) => {
      const text = line.trim();
      const inline = SECTION_HEADING.exec(text);
      const section = inline ? inline[1] : SECTION_NUMBER.exec(text)?.[1];
      if (!section || found.has(section)) return;
      const title = inline ? inline[2] : titleBelow(lines, at);
      if (!title || title.includes(LEADER)) return;
      found.set(section, { section, page: index + 1, title });
    });
  });
  return [...found.values()];
}

/** Section numbers sort by their parts as numbers, so 17.18.5 stands before 17.18.49 */
function bySection(a, b) {
  const left = a.section.split(".").map(Number);
  const right = b.section.split(".").map(Number);
  for (let i = 0; i < Math.max(left.length, right.length); i += 1) {
    const diff = (left[i] ?? -1) - (right[i] ?? -1);
    if (diff !== 0) return diff;
  }
  return 0;
}

async function build() {
  const rows = [];
  for (const { part, file } of PARTS) {
    if (await missing(join(pdfDir, file))) {
      throw new Error(
        `${basename(file)} is not in spec/pdf/. Run scripts/fetch-ooxml-spec.sh first.`
      );
    }
    process.stdout.write(`  reading ${part}\n`);
    const headings = headingsOf(await pageTexts(file)).sort(bySection);
    for (const { section, page, title } of headings) {
      rows.push(`${part}\t${section}\t${page}\t${title}`);
    }
    process.stdout.write(`  ${headings.length} sections in ${part}\n`);
  }
  await writeFile(indexFile, `${rows.join("\n")}\n`, "utf8");
  process.stdout.write(
    `\nwrote spec/pdf/section-index.tsv, ${rows.length} sections\n`
  );
  return rows;
}

async function rowsForLookup() {
  if (await missing(indexFile)) return build();
  const text = await readFile(indexFile, "utf8");
  return text.split("\n").filter(Boolean);
}

/** A query is a section number when it is digits and dots, and a title search otherwise */
function matches(rows, query) {
  const byNumber = /^\d+(\.\d+)*$/.test(query);
  const needle = query.toLowerCase();
  return rows.filter((row) => {
    const [, section, , title] = row.split("\t");
    return byNumber
      ? section === query || section.startsWith(`${query}.`)
      : title.toLowerCase().includes(needle);
  });
}

// A lookup is read through `head` and `grep` often enough that the closed pipe has to be survivable
process.stdout.on("error", (error) => {
  if (error.code !== "EPIPE") throw error;
});

const query = process.argv[2];
if (!query) {
  await build();
} else {
  const found = matches(await rowsForLookup(), query);
  if (found.length === 0) {
    process.stdout.write(`nothing indexed for "${query}"\n`);
    process.exitCode = 1;
  }
  for (const row of found) {
    const [part, section, page, title] = row.split("\t");
    process.stdout.write(`${part} p.${page}  §${section}  ${title}\n`);
  }
}
