// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { describe, expect, expectTypeOf, it } from "vitest";
import { decode, makeDocx } from "../__testing__/docx";
import { rangeOfText } from "../__testing__/editing";
import {
  addComment,
  setCommentResolved,
} from "../editor/commands/commentCommands";
import { createEditorState } from "../editor/createEditor";
import { decodeUtf8, elementChildren, parseXml } from "../ooxml/xml";
import type {
  CommentOnlyVerdict,
  onlyCommentsChangedBy,
} from "./commentOnlyChange";
import { commentsPlanner } from "./comments/writing";
import { importDocx } from "./importDocx";
import { CONTENT_TYPES_PATH, contentTypeWriter } from "./packageParts";
import {
  type ChangeVerdict,
  type PackageReason,
  partsKept,
  protectionPolicyFor,
} from "./protectionPolicy";
import type { NewRelationship, RelationshipWriter } from "./relationships";

const encoder = new TextEncoder();
const COMMENTS_PART = "word/comments.xml";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A document with a comment part to gain, and a content types part to declare it in */
function plainDocx(): Uint8Array {
  const parts = unzipSync(makeDocx(`<w:p>${run("Alpha beta")}</w:p>`));
  parts[CONTENT_TYPES_PATH] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

function applied(state: EditorState, command: Command): EditorState {
  let next = state;
  expect(command(state, (tr) => (next = state.apply(tr)))).toBe(true);
  return next;
}

/** The document with a comment of this author's on "beta", settled so all three parts are written */
function commentedAndSettled(doc: EditorState["doc"]): EditorState {
  const opened = createEditorState(doc);
  const { from, to } = rangeOfText(opened.doc, "beta");
  const selected = opened.apply(
    opened.tr.setSelection(TextSelection.create(opened.doc, from, to))
  );
  const added = applied(
    selected,
    addComment({ text: "note", author: "Someone", authorId: "me" })
  );
  return applied(added, setCommentResolved("0", true));
}

/** A relationship writer that records what a planner asks it for */
function recordingWriter(): {
  added: NewRelationship[];
  writer: RelationshipWriter;
} {
  const added: NewRelationship[] = [];
  return {
    added,
    writer: {
      opened: [],
      add: (relationship) => {
        added.push(relationship);
        return `rId${added.length}`;
      },
      part: () => null,
    },
  };
}

/** Every content type the part declares an Override for */
function overridesIn(bytes: Uint8Array | undefined): string[] {
  if (bytes === undefined) return [];
  return elementChildren(parseXml(decodeUtf8(bytes).text).documentElement)
    .filter((el) => el.localName === "Override")
    .map((el) => el.getAttribute("ContentType") ?? "");
}

const sorted = (values: Iterable<string>) => Array.from(values).sort();

/**
 * The policy, loaded the way the registry needs it to be loaded.
 *
 * Nothing else in this file reaches `./comments/policy`: the declaration `commentOnlyChange`
 * carries is imported as a type, which is erased before anything runs.
 */
const commentsPolicy = async () =>
  (await import("./comments/policy")).commentsPolicy;

describe("the protection policy registry", () => {
  /**
   * The registry is a map filled as modules are evaluated, so a level answers only where
   * something has imported the module that declares it. `commentOnlyChange` is what imports the
   * comments policy for a caller of the public entry, and the line below is what imports it here.
   * A second policy is invisible to `protectionPolicyFor` until a module in the graph reaches it
   * the same way: adding one means finding it an importer, not only writing it.
   */
  it("answers for a level only where the module declaring it was loaded", async () => {
    expect(protectionPolicyFor("none")).toBeUndefined();
    expect(protectionPolicyFor("readOnly")).toBeUndefined();

    const policy = await commentsPolicy();

    expect(protectionPolicyFor("comments")).toBe(policy);
    expect(policy.level).toBe("comments");
  });
});

describe("the comments policy and the comment part planners", () => {
  /**
   * The planners write the three parts and the verifier excuses them from the byte comparison,
   * and both sides read the policy for which parts those are. A relationship or a content type
   * the planners wrote that the policy did not name would be a part the writer adds and the
   * verifier refuses; one the policy named that no planner writes would be a part a submission
   * may bring that this editor never writes.
   */
  it("every relationship and content type the comment planners add is one the comments policy allows", async () => {
    const bytes = plainDocx();
    const { doc, session } = importDocx(bytes);
    const state = commentedAndSettled(doc);
    const { added, writer } = recordingWriter();

    const contentTypes = contentTypeWriter(session.parts);
    const planned = commentsPlanner.plan(state.doc, session, {
      relationships: writer,
      contentTypes,
    });
    if (planned === null) throw new Error("the planners wrote nothing");

    const policy = await commentsPolicy();
    const declared = overridesIn(contentTypes.part() ?? undefined);
    const before = overridesIn(session.parts.get(CONTENT_TYPES_PATH));

    expect(sorted(added.map((relationship) => relationship.type))).toEqual(
      sorted(policy.parts.map((kind) => kind.relType))
    );
    expect(sorted(declared.filter((type) => !before.includes(type)))).toEqual(
      sorted(policy.parts.map((kind) => kind.contentType))
    );
  });

  it("a part kind names one path for a package that has it and one for a package that does not", async () => {
    const policy = await commentsPolicy();
    const [comments] = policy.parts;
    const bytes = plainDocx();
    const empty = importDocx(bytes).session;

    expect(comments.pathIn(empty)).toBeNull();
    expect(comments.writePathIn(empty)).toBe(COMMENTS_PART);

    // A package already holding a part of some other making at that name gets the next one
    const taken = importDocx(
      zipSync({
        ...unzipSync(bytes),
        [COMMENTS_PART]: encoder.encode("<w:comments/>"),
      })
    ).session;

    expect(taken.parts.has(COMMENTS_PART)).toBe(true);
    expect(comments.pathIn(taken)).toBeNull();
    expect(comments.writePathIn(taken)).toBe("word/comments2.xml");
  });

  it("answers with the policy's own reason for a rewritten entry that is not well formed or not allowed", async () => {
    const policy = await commentsPolicy();
    const bytes = plainDocx();
    const { doc, session } = importDocx(bytes);
    const state = commentedAndSettled(doc);
    const { exportDocx } = await import("./exportDocx");
    const commented = exportDocx(state.doc, session);
    const text = decode(unzipSync(commented)[COMMENTS_PART]);
    const forged = zipSync({
      ...unzipSync(commented),
      [COMMENTS_PART]: encoder.encode(
        text.replace(
          "</w:p></w:comment>",
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p></w:comment>'
        )
      ),
    });

    expect(
      partsKept(policy, importDocx(commented), importDocx(forged), "me", {
        editableComments: "own",
      })
    ).toEqual({
      ok: false,
      reason: policy.rejectedMarkup,
      part: COMMENTS_PART,
    });
  });

  /**
   * The public declaration and the shape the policy answers in are the same type without an alias
   * standing between them, so widening one and not the other stops compiling here.
   */
  it("answers in the very type the public entry declares", () => {
    expectTypeOf<CommentOnlyVerdict>().toEqualTypeOf<
      ReturnType<typeof onlyCommentsChangedBy>
    >();
    expectTypeOf<CommentOnlyVerdict>().toEqualTypeOf<
      ChangeVerdict<
        "body-changed" | "comment-not-owned" | "comment-author-forged",
        "comment-markup-rejected" | PackageReason
      >
    >();
  });
});
