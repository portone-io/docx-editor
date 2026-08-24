// @vitest-environment jsdom
import { unzipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  anchoredDrawingXml,
  bytesEqual,
  decode,
  decodeBase64,
  drawingRun,
  exportErrorCode,
  IMAGE_REL_ID,
  inlineDrawingXml,
  makeDocx,
  makeImageDocx,
  OTHER_PNG_BASE64,
  TINY_PNG,
  TINY_PNG_BASE64,
  TINY_PNG_DATA_URL,
} from "../__testing__/docx";
import { type ImageExtent, imageBase64Of } from "../ooxml/image";
import { docxSchema } from "../schema";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import { MAX_IMAGE_BYTES } from "./media";

const SEAL = `<w:p><w:r><w:t xml:space="preserve">Signature</w:t></w:r>${drawingRun(
  inlineDrawingXml({ descr: "Seal" })
)}</w:p>`;

function imagesIn(doc: PMNode): PMNode[] {
  const found: PMNode[] = [];
  doc.descendants((node) => {
    if (node.type.name === "image") found.push(node);
    return true;
  });
  return found;
}

/** The bytes an image node is carrying around */
function pixelsOf(node: PMNode): Uint8Array {
  const src: unknown = node.attrs.src;
  if (typeof src !== "string") throw new Error("the image has no src");
  return decodeBase64(imageBase64Of(src));
}

function replaceInline(
  doc: PMNode,
  replace: (child: PMNode) => PMNode
): PMNode {
  const blocks = doc.children.map((block) =>
    block.type.name === "paragraph"
      ? block.copy(Fragment.from(block.children.map(replace)))
      : block
  );
  return docxSchema.nodes.doc.create(null, blocks);
}

function resizeImages(doc: PMNode, extent: ImageExtent): PMNode {
  return replaceInline(doc, (child) =>
    child.type.name === "image"
      ? docxSchema.nodes.image.create(
          { ...child.attrs, extent },
          null,
          child.marks
        )
      : child
  );
}

/** The document with a freshly created image node appended to the first paragraph */
function insertImage(doc: PMNode, attrs: Record<string, unknown>): PMNode {
  const first = doc.child(0);
  const withImage = first.copy(
    Fragment.from([...first.children, docxSchema.nodes.image.create(attrs)])
  );
  const blocks = doc.children.map((block, index) =>
    index === 0 ? withImage : block
  );
  return docxSchema.nodes.doc.create(null, blocks);
}

function partsOf(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipSync(bytes);
}

/** Every part except the body is exactly as it was, and no part came or went */
function expectOnlyBodyRewritten(
  original: Uint8Array,
  exported: Uint8Array,
  mainPartPath: string
): void {
  const before = partsOf(original);
  const after = partsOf(exported);
  expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
  for (const path of Object.keys(before)) {
    if (path === mainPartPath) continue;
    expect(bytesEqual(after[path], before[path])).toBe(true);
  }
}

describe("opening a document with an image", () => {
  const { doc } = importDocx(makeImageDocx(SEAL));

  it("reads the picture as an image node carrying its bytes", () => {
    const [image] = imagesIn(doc);

    expect(image.attrs.src).toBe(TINY_PNG_DATA_URL);
    expect(image.attrs.extent).toEqual({ cx: 1905000, cy: 952500 });
    expect(image.attrs.alt).toBe("Seal");
    expect(bytesEqual(pixelsOf(image), TINY_PNG)).toBe(true);
  });

  it("keeps the original drawing XML on the node", () => {
    expect(imagesIn(doc)[0].attrs.xml).toContain("<pic:pic>");
  });

  it("keeps the run formatting the picture sat in", () => {
    const mark = imagesIn(doc)[0].marks.find(
      (entry) => entry.type.name === "run"
    );
    expect(mark?.attrs.rPr).toBe("<w:rPr><w:noProof/></w:rPr>");
  });

  it("leaves the paragraph editable around the image", () => {
    expect(doc.child(0).type.name).toBe("paragraph");
    expect(doc.child(0).textContent).toBe("Signature");
  });
});

describe("a drawing we do not interpret", () => {
  const cases = {
    floating: anchoredDrawingXml(),
    "not a picture": inlineDrawingXml({
      uri: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    }),
    "linked to a file outside the package": inlineDrawingXml({ linked: true }),
    "pointing at a relationship that is not there": inlineDrawingXml({
      relId: "rIdGone",
    }),
  };

  for (const [name, drawing] of Object.entries(cases)) {
    it(`${name}: keeps the whole paragraph as preserved content`, () => {
      const bytes = makeImageDocx(`<w:p>${drawingRun(drawing)}</w:p>`);
      const { doc } = importDocx(bytes);
      expect(doc.child(0).type.name).toBe("docxRaw");
      expect(imagesIn(doc)).toEqual([]);
    });

    it(`${name}: still round trips byte for byte`, () => {
      const bytes = makeImageDocx(`<w:p>${drawingRun(drawing)}</w:p>`);
      const { doc, session } = importDocx(bytes);
      const out = exportDocx(doc, session);
      expectOnlyBodyRewritten(bytes, out, session.mainPartPath);
      expect(decode(partsOf(out)[session.mainPartPath])).toBe(
        decode(partsOf(bytes)[session.mainPartPath])
      );
    });
  }
});

/**
 * The session holds every image it draws as base64 for as long as the editor is open, so
 * an image past the cap takes the same road as one no browser can draw
 */
describe("an image weighing more than we hold", () => {
  const oversized = `<w:p>${drawingRun(inlineDrawingXml())}</w:p>`;
  const makeOversized = () =>
    makeImageDocx(oversized, {
      bytes: new Uint8Array(MAX_IMAGE_BYTES + 1),
    });

  it("keeps the whole paragraph as preserved content", () => {
    const { doc } = importDocx(makeOversized());
    expect(doc.child(0).type.name).toBe("docxRaw");
    expect(imagesIn(doc)).toEqual([]);
  });

  it("still round trips byte for byte", () => {
    const bytes = makeOversized();
    const { doc, session } = importDocx(bytes);
    const out = exportDocx(doc, session);
    expectOnlyBodyRewritten(bytes, out, session.mainPartPath);
  });
});

describe("exporting a document with an image nobody touched", () => {
  it("is byte identical, media part and all", () => {
    const bytes = makeImageDocx(SEAL);
    const { doc, session } = importDocx(bytes);
    const out = exportDocx(doc, session);

    expectOnlyBodyRewritten(bytes, out, session.mainPartPath);
    expect(decode(partsOf(out)[session.mainPartPath])).toBe(
      decode(partsOf(bytes)[session.mainPartPath])
    );
  });

  it("writes the drawing back as it was even when the paragraph around it changed", () => {
    const bytes = makeImageDocx(SEAL);
    const { doc, session } = importDocx(bytes);
    const edited = replaceInline(doc, (child) =>
      child.isText ? docxSchema.text("edited signature", child.marks) : child
    );
    const out = exportDocx(edited, session);
    const documentXml = decode(partsOf(out)[session.mainPartPath]);

    expect(documentXml).toContain("edited signature");
    expect(documentXml).toContain(
      drawingRun(inlineDrawingXml({ descr: "Seal" }))
    );
    expectOnlyBodyRewritten(bytes, out, session.mainPartPath);
  });

  it("does not put a data URL anywhere near the export", () => {
    const bytes = makeImageDocx(SEAL);
    const { doc, session } = importDocx(bytes);
    expect(
      decode(partsOf(exportDocx(doc, session))[session.mainPartPath])
    ).not.toContain(TINY_PNG_BASE64);
  });
});

describe("resizing an imported image", () => {
  const bytes = makeImageDocx(SEAL);

  it("rewrites only the extents and leaves the media part alone", () => {
    const { doc, session } = importDocx(bytes);
    const out = exportDocx(
      resizeImages(doc, { cx: 952500, cy: 476250 }),
      session
    );
    const documentXml = decode(partsOf(out)[session.mainPartPath]);

    expect(documentXml).toContain('<wp:extent cx="952500" cy="476250"/>');
    expect(documentXml).toContain('<a:ext cx="952500" cy="476250"/>');
    expect(documentXml).not.toContain("1905000");
    expect(documentXml).toContain(`r:embed="${IMAGE_REL_ID}"`);
    expectOnlyBodyRewritten(bytes, out, session.mainPartPath);
  });

  it("reads the new size back on reopening", () => {
    const { doc, session } = importDocx(bytes);
    const out = exportDocx(
      resizeImages(doc, { cx: 952500, cy: 476250 }),
      session
    );
    const [image] = imagesIn(importDocx(out).doc);

    expect(image.attrs.extent).toEqual({ cx: 952500, cy: 476250 });
    expect(bytesEqual(pixelsOf(image), TINY_PNG)).toBe(true);
  });
});

const OTHER_PNG_DATA_URL = `data:image/png;base64,${OTHER_PNG_BASE64}`;
const INSERTED = {
  src: OTHER_PNG_DATA_URL,
  extent: { cx: 952500, cy: 952500 },
  alt: "a new picture",
  xml: null,
};

describe("inserting an image that was not in the document", () => {
  const bytes = makeImageDocx(SEAL);
  const opened = () => importDocx(bytes);

  it("adds a media part holding exactly the bytes the node carried", () => {
    const { doc, session } = opened();
    const out = exportDocx(insertImage(doc, INSERTED), session);
    const parts = partsOf(out);

    const added = Object.keys(parts).filter(
      (path) =>
        path.startsWith("word/media/") && path !== "word/media/image1.png"
    );
    expect(added.length).toBe(1);
    expect(bytesEqual(parts[added[0]], decodeBase64(OTHER_PNG_BASE64))).toBe(
      true
    );
    // The image that was already there is untouched
    expect(bytesEqual(parts["word/media/image1.png"], TINY_PNG)).toBe(true);
  });

  it("relates the new part to the body without disturbing the relationships already there", () => {
    const { doc, session } = opened();
    const out = exportDocx(insertImage(doc, INSERTED), session);
    const rels = decode(partsOf(out)["word/_rels/document.xml.rels"]);
    const documentXml = decode(partsOf(out)[session.mainPartPath]);

    expect(rels).toContain(`Id="${IMAGE_REL_ID}"`);
    const added = Array.from(rels.matchAll(/Id="([^"]+)"/g)).map(
      (match) => match[1]
    );
    expect(added).toHaveLength(2);
    const relId = added.find((id) => id !== IMAGE_REL_ID);
    expect(relId).toBeDefined();
    expect(documentXml).toContain(`r:embed="${relId}"`);
    expect(rels).toContain('Target="media/image-');
  });

  it("declares the content type the new part needs", () => {
    const { doc, session } = opened();
    const out = exportDocx(insertImage(doc, INSERTED), session);
    const types = decode(partsOf(out)["[Content_Types].xml"]);

    expect(types).toContain(
      '<Default Extension="png" ContentType="image/png"/>'
    );
    expect(types).toContain('PartName="/word/document.xml"');
  });

  it("leaves the content types alone when the kind is already declared", () => {
    const declared = makeImageDocx(SEAL, { declarePng: true });
    const { doc, session } = importDocx(declared);
    const out = exportDocx(insertImage(doc, INSERTED), session);

    expect(
      bytesEqual(
        partsOf(out)["[Content_Types].xml"],
        partsOf(declared)["[Content_Types].xml"]
      )
    ).toBe(true);
  });

  it("comes back as an image with the same pixels when the export is reopened", () => {
    const { doc, session } = opened();
    const out = exportDocx(insertImage(doc, INSERTED), session);
    const reopened = importDocx(out);
    const inserted = imagesIn(reopened.doc)[1];

    expect(inserted.attrs.src).toBe(OTHER_PNG_DATA_URL);
    expect(inserted.attrs.extent).toEqual({ cx: 952500, cy: 952500 });
    expect(inserted.attrs.alt).toBe("a new picture");
    expect(bytesEqual(pixelsOf(inserted), decodeBase64(OTHER_PNG_BASE64))).toBe(
      true
    );
    // Exporting the reopened document changes nothing further
    expect(bytesEqual(exportDocx(reopened.doc, reopened.session), out)).toBe(
      true
    );
  });

  it("reuses the part already in the package for bytes it already holds", () => {
    const { doc, session } = opened();
    const out = exportDocx(
      insertImage(doc, { ...INSERTED, src: TINY_PNG_DATA_URL }),
      session
    );
    const parts = partsOf(out);

    expect(Object.keys(parts).sort()).toEqual(
      Object.keys(partsOf(bytes)).sort()
    );
    expect(
      decode(parts[session.mainPartPath]).match(
        new RegExp(`r:embed="${IMAGE_REL_ID}"`, "g")
      )
    ).toHaveLength(2);
  });

  it("writes one part for the same image inserted twice", () => {
    const { doc, session } = opened();
    const twice = insertImage(insertImage(doc, INSERTED), INSERTED);
    const parts = partsOf(exportDocx(twice, session));

    expect(
      Object.keys(parts).filter((path) => path.startsWith("word/media/"))
    ).toHaveLength(2);
  });

  it("gives the two drawings ids of their own, past the ones the document uses", () => {
    const { doc, session } = opened();
    const twice = insertImage(insertImage(doc, INSERTED), INSERTED);
    const documentXml = decode(
      partsOf(exportDocx(twice, session))[session.mainPartPath]
    );
    const ids = Array.from(documentXml.matchAll(/<wp:docPr id="(\d+)"/g)).map(
      (match) => Number(match[1])
    );

    // The imported drawing states 4, so nothing new may claim 4 or below
    expect(ids).toEqual([4, 5, 6]);
  });
});

const TABLE_WITH_IMAGE =
  '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  `<w:tr><w:tc><w:p>${drawingRun(inlineDrawingXml())}</w:p></w:tc></w:tr></w:tbl>`;

describe("an image inside a table cell", () => {
  const bytes = makeImageDocx(TABLE_WITH_IMAGE);

  it("opens as an image node in the cell and round trips byte for byte", () => {
    const { doc, session } = importDocx(bytes);
    expect(doc.child(0).type.name).toBe("table");
    expect(imagesIn(doc)).toHaveLength(1);
    expect(imagesIn(doc)[0].attrs.src).toBe(TINY_PNG_DATA_URL);

    const out = exportDocx(doc, session);
    expectOnlyBodyRewritten(bytes, out, session.mainPartPath);
    expect(decode(partsOf(out)[session.mainPartPath])).toBe(
      decode(partsOf(bytes)[session.mainPartPath])
    );
  });

  it("takes a newly inserted image alongside the one already there", () => {
    const { doc, session } = importDocx(bytes);
    const table = doc.child(0);
    const row = table.child(0);
    const cell = row.child(0);
    const paragraph = cell.child(0);
    const edited = docxSchema.nodes.doc.create(null, [
      table.copy(
        Fragment.from([
          row.copy(
            Fragment.from([
              cell.copy(
                Fragment.from([
                  paragraph.copy(
                    Fragment.from([
                      ...paragraph.children,
                      docxSchema.nodes.image.create(INSERTED),
                    ])
                  ),
                ])
              ),
            ])
          ),
        ])
      ),
    ]);
    const out = exportDocx(edited, session);
    const parts = partsOf(out);

    expect(
      decode(parts[session.mainPartPath]).match(/<w:drawing/g)
    ).toHaveLength(2);
    expect(
      Object.keys(parts).filter((path) => path.startsWith("word/media/"))
    ).toHaveLength(2);
    expect(imagesIn(importDocx(out).doc)).toHaveLength(2);
  });
});

describe("an image the export cannot write", () => {
  it("refuses a package with no content types to declare the new part in", () => {
    const { doc, session } = importDocx(
      makeDocx('<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>')
    );
    expect(
      exportErrorCode(() => exportDocx(insertImage(doc, INSERTED), session))
    ).toBe("missing-content-types");
  });

  it("refuses an image that carries neither original XML nor bytes", () => {
    const { doc, session } = importDocx(makeImageDocx(SEAL));
    const empty = insertImage(doc, {
      src: null,
      extent: null,
      alt: null,
      xml: null,
    });
    expect(exportErrorCode(() => exportDocx(empty, session))).toBe(
      "lost-original"
    );
  });
});

describe("a document with no image at all", () => {
  it("gains no media part, no relationship and no content type", () => {
    const bytes = makeImageDocx(
      '<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>'
    );
    const { doc, session } = importDocx(bytes);
    const edited = replaceInline(doc, (child) =>
      child.isText ? docxSchema.text("edited body", child.marks) : child
    );
    const out = exportDocx(edited, session);

    expect(decode(partsOf(out)[session.mainPartPath])).toContain("edited body");
    expectOnlyBodyRewritten(bytes, out, session.mainPartPath);
  });
});
