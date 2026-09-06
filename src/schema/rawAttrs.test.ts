// @vitest-environment jsdom
/**
 * Every raw fragment the schema draws, held in both directions at once.
 *
 * Thirteen of the attrs that carry one were read through the gate without a single test noticing
 * if they stopped being: replacing the call with a bare `dom.getAttribute` failed nothing. The
 * table below names every attr the writer writes from, so a new one has to say whether it carries
 * a fragment before this file will compile past `attrRoles`, and each fragment it does name is
 * drawn twice: once holding its shape, which has to come back, and once not, which has to be gone.
 */
import {
  DOMSerializer,
  Fragment,
  DOMParser as PMDOMParser,
  type Node as PMNode,
} from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { type AttrRole, MARK_ATTR_ROLES, NODE_ATTR_ROLES } from "./attrRoles";
import { docxSchema } from "./index";

const serializer = DOMSerializer.fromSchema(docxSchema);
const parser = PMDOMParser.fromSchema(docxSchema);

/** A 1x1 png, the shortest src an image node can carry */
const PNG_SRC =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAf" +
  "FcSJAAAAC0lEQVR42mNgAAIAAAUAAen63NgAAAAASUVORK5CYII=";

interface RawAttrCase {
  /** The `data-` attribute the fragment is drawn as */
  readonly attribute: string;
  /** A fragment holding the shape this attr goes back out as */
  readonly sound: string;
  /** One that does not hold it, spliced into the exported file if the rule let it through */
  readonly adversarial: string;
  /** The block that carries the fragment, drawn around whatever the attr sits on */
  readonly draw: (fragment: string) => PMNode;
}

/** null for a source attr carrying no XML, which no shape has anything to say about */
type RawAttr = RawAttrCase | null;

type RawAttrTable = Readonly<Record<string, Readonly<Record<string, RawAttr>>>>;

function paragraph(attrs: Record<string, unknown>, inline: PMNode[]): PMNode {
  return docxSchema.nodes.paragraph.create(attrs, inline);
}

function cell(attrs: Record<string, unknown>): PMNode {
  return docxSchema.nodes.tableCell.create(attrs, [
    paragraph({}, [docxSchema.text("x")]),
  ]);
}

function row(attrs: Record<string, unknown>): PMNode {
  return docxSchema.nodes.tableRow.create(attrs, [cell({})]);
}

function table(attrs: Record<string, unknown>, rows: PMNode[]): PMNode {
  return docxSchema.nodes.table.create(attrs, rows);
}

function inline(node: PMNode): PMNode {
  return paragraph({}, [node]);
}

function marked(name: string, attrs: Record<string, unknown>): PMNode {
  return paragraph({}, [
    docxSchema.text("x", [docxSchema.marks[name].create(attrs)]),
  ]);
}

/** A list that ends the tag it was written into and opens a run of its own beside the content */
const SMUGGLING_LIST = 'w:rsidR="00A"><w:r><w:t>smuggled</w:t></w:r';

/** An element with a run standing beside it, which the writer would splice in along with it */
const withSibling = (element: string) =>
  `${element}<w:r><w:t>smuggled</w:t></w:r>`;

const NODE_FRAGMENTS: RawAttrTable = {
  paragraph: {
    pAttrs: {
      attribute: "data-pattrs",
      sound: 'w:rsidR="00A1B2C3"',
      adversarial: SMUGGLING_LIST,
      draw: (xml) => paragraph({ pAttrs: xml }, [docxSchema.text("x")]),
    },
    pPr: {
      attribute: "data-ppr",
      sound: '<w:pPr><w:jc w:val="center"/></w:pPr>',
      adversarial: withSibling("<w:pPr/>"),
      draw: (xml) => paragraph({ pPr: xml }, [docxSchema.text("x")]),
    },
  },
  table: {
    tblAttrs: {
      attribute: "data-tblattrs",
      sound: 'w:rsidTr="00A1B2C3"',
      adversarial: SMUGGLING_LIST,
      draw: (xml) => table({ tblAttrs: xml }, [row({})]),
    },
    tblPr: {
      attribute: "data-tblpr",
      sound: '<w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>',
      adversarial: withSibling("<w:tblPr/>"),
      draw: (xml) => table({ tblPr: xml }, [row({})]),
    },
    tblW: null,
    gridCols: null,
  },
  tableRow: {
    trAttrs: {
      attribute: "data-trattrs",
      sound: 'w:rsidR="00B2C3D4"',
      adversarial: SMUGGLING_LIST,
      draw: (xml) => table({}, [row({ trAttrs: xml })]),
    },
    tblPrEx: {
      attribute: "data-tblprex",
      sound: '<w:tblPrEx><w:tblLayout w:type="fixed"/></w:tblPrEx>',
      adversarial: withSibling("<w:tblPrEx/>"),
      draw: (xml) => table({}, [row({ tblPrEx: xml })]),
    },
    trPr: {
      attribute: "data-trpr",
      sound: "<w:trPr><w:cantSplit/></w:trPr>",
      adversarial: withSibling("<w:trPr/>"),
      draw: (xml) => table({}, [row({ trPr: xml })]),
    },
  },
  tableCell: {
    tcAttrs: {
      attribute: "data-tcattrs",
      sound: 'w:rsidR="00C3D4E5"',
      adversarial: SMUGGLING_LIST,
      draw: (xml) =>
        table({}, [
          docxSchema.nodes.tableRow.create(null, [cell({ tcAttrs: xml })]),
        ]),
    },
    tcPr: {
      attribute: "data-tcpr",
      sound: '<w:tcPr><w:vAlign w:val="center"/></w:tcPr>',
      adversarial: withSibling("<w:tcPr/>"),
      draw: (xml) =>
        table({}, [
          docxSchema.nodes.tableRow.create(null, [cell({ tcPr: xml })]),
        ]),
    },
    sdtPrefix: {
      attribute: "data-sdt-prefix",
      sound: '<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr>',
      adversarial:
        "<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>smuggled</w:t></w:r></w:sdtContent>",
      draw: (xml) =>
        table({}, [
          docxSchema.nodes.tableRow.create(null, [cell({ sdtPrefix: xml })]),
        ]),
    },
    colspan: null,
    rowspan: null,
    tcW: null,
    sdtContentsLocked: null,
    sdtDeletionLocked: null,
  },
  rawBlock: {
    xml: {
      attribute: "data-xml",
      sound: "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>",
      adversarial: withSibling("<w:tbl/>"),
      draw: (xml) => docxSchema.nodes.rawBlock.create({ xml, name: "w:tbl" }),
    },
    name: null,
  },
  docxRaw: {},
  bookmarkBlock: {},
  hardBreak: {
    brAttrs: {
      attribute: "data-battrs",
      sound: 'w:type="page"',
      adversarial: SMUGGLING_LIST,
      draw: (xml) =>
        inline(docxSchema.nodes.hardBreak.create({ brAttrs: xml })),
    },
  },
  image: {
    xml: {
      attribute: "data-xml",
      sound: "<w:drawing><wp:inline/></w:drawing>",
      adversarial: withSibling("<w:drawing/>"),
      draw: (xml) =>
        inline(docxSchema.nodes.image.create({ src: PNG_SRC, xml })),
    },
    src: null,
    extent: null,
    alt: null,
  },
  commentStart: {
    xml: {
      attribute: "data-xml",
      sound: '<w:commentRangeStart w:id="7"/>',
      adversarial: withSibling('<w:commentRangeStart w:id="7"/>'),
      draw: (xml) =>
        inline(docxSchema.nodes.commentStart.create({ id: "7", xml })),
    },
    id: null,
  },
  commentEnd: {
    xml: {
      attribute: "data-xml",
      sound: '<w:commentRangeEnd w:id="7"/>',
      adversarial: withSibling('<w:commentRangeEnd w:id="7"/>'),
      draw: (xml) =>
        inline(docxSchema.nodes.commentEnd.create({ id: "7", xml })),
    },
    id: null,
  },
  commentReference: {
    referenceXml: {
      attribute: "data-reference-xml",
      sound: '<w:commentReference w:id="7"/>',
      adversarial: withSibling('<w:commentReference w:id="7"/>'),
      draw: (xml) =>
        inline(
          docxSchema.nodes.commentReference.create({
            id: "7",
            referenceXml: xml,
          })
        ),
    },
    commentXml: {
      attribute: "data-comment-xml",
      sound: '<w:comment w:id="7"><w:p/></w:comment>',
      adversarial: '<w:comment w:id="7"/><w:comment w:id="8"/>',
      draw: (xml) =>
        inline(
          docxSchema.nodes.commentReference.create({ id: "7", commentXml: xml })
        ),
    },
    extensionXml: {
      attribute: "data-comment-extension-xml",
      sound: '<w15:commentEx w15:paraId="0A0A0A0A" w15:done="0"/>',
      adversarial: withSibling('<w15:commentEx w15:paraId="0A0A0A0A"/>'),
      draw: (xml) =>
        inline(
          docxSchema.nodes.commentReference.create({
            id: "7",
            extensionXml: xml,
          })
        ),
    },
    // A reply's body goes out into `word/comments.xml` rather than into the story, so it is held
    // to its shape there. Its extended properties travel the same way, held below
    replies: {
      attribute: "data-comment-replies",
      sound: '<w:comment w:id="8"><w:p/></w:comment>',
      adversarial: '<w:comment w:id="8"/><w:comment w:id="9"/>',
      draw: (xml) =>
        inline(
          docxSchema.nodes.commentReference.create({
            id: "7",
            replies: [
              {
                id: "8",
                paraId: "0B0B0B0B",
                parentParaId: "0A0A0A0A",
                commentXml: xml,
              },
            ],
          })
        ),
    },
    id: null,
    author: null,
    authorId: null,
    initials: null,
    date: null,
    text: null,
    paraId: null,
    resolved: null,
  },
  noteReference: {
    referenceXml: {
      attribute: "data-reference-xml",
      sound: '<w:footnoteReference w:id="2"/>',
      adversarial: withSibling('<w:footnoteReference w:id="2"/>'),
      draw: (xml) =>
        inline(
          docxSchema.nodes.noteReference.create({ id: "2", referenceXml: xml })
        ),
    },
    kind: null,
    id: null,
    customMarkFollows: null,
  },
  rawInline: {
    xml: {
      attribute: "data-xml",
      sound: '<w:bookmarkStart w:id="1" w:name="clause"/>',
      adversarial: withSibling('<w:bookmarkStart w:id="1" w:name="clause"/>'),
      draw: (xml) => inline(docxSchema.nodes.rawInline.create({ xml })),
    },
  },
};

const MARK_FRAGMENTS: RawAttrTable = {
  run: {
    rAttrs: {
      attribute: "data-rattrs",
      sound: 'w:rsidR="00D4E5F6"',
      adversarial: SMUGGLING_LIST,
      draw: (xml) => marked("run", { rAttrs: xml }),
    },
    rPr: {
      attribute: "data-rpr",
      sound: "<w:rPr><w:b/></w:rPr>",
      adversarial: withSibling("<w:rPr/>"),
      draw: (xml) => marked("run", { rPr: xml }),
    },
  },
  sdt: {
    sdtPrefix: {
      attribute: "data-sdt-prefix",
      sound: '<w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr>',
      adversarial:
        "<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>smuggled</w:t></w:r></w:sdtContent>",
      draw: (xml) => marked("sdt", { sdtPrefix: xml }),
    },
    contentsLocked: null,
    deletionLocked: null,
  },
  link: {
    linkPrefix: {
      attribute: "data-link-prefix",
      sound: '<w:hyperlink r:id="rId7" w:history="1">',
      adversarial: '<w:hyperlink r:id="rId7"><w:r><w:t>smuggled</w:t></w:r>',
      draw: (xml) =>
        marked("link", { linkPrefix: xml, href: "https://example.com" }),
    },
    href: null,
  },
  tab: {
    tabAttrs: {
      attribute: "data-tattrs",
      sound: 'w:pos="720"',
      adversarial: SMUGGLING_LIST,
      draw: (xml) =>
        paragraph({}, [
          docxSchema.text("\t", [
            docxSchema.marks.tab.create({ tabAttrs: xml }),
          ]),
        ]),
    },
  },
};

function render(block: PMNode): HTMLElement {
  const host = document.createElement("div");
  host.appendChild(serializer.serializeFragment(Fragment.from(block)));
  return host;
}

function roundTrip(block: PMNode): PMNode {
  return parser.parse(render(block), { preserveWhitespace: true });
}

/** Every string any node or mark of the tree ended up carrying, JSON attrs walked into */
function stringsIn(value: unknown): string[] {
  if (typeof value === "string") return [value];
  if (Array.isArray(value)) return value.flatMap(stringsIn);
  if (typeof value === "object" && value !== null) {
    return Object.values(value).flatMap(stringsIn);
  }
  return [];
}

function attrStrings(doc: PMNode): string[] {
  const found: string[] = [];
  doc.descendants((node) => {
    found.push(...stringsIn(node.attrs));
    for (const mark of node.marks) found.push(...stringsIn(mark.attrs));
    return true;
  });
  return found;
}

function sourceAttrs(roles: Readonly<Record<string, AttrRole>>): string[] {
  return Object.entries(roles)
    .filter(([, role]) => role === "source")
    .map(([name]) => name)
    .sort();
}

function casesOf(table: RawAttrTable): [string, string, RawAttrCase][] {
  return Object.entries(table).flatMap(([type, attrs]) =>
    Object.entries(attrs).flatMap<[string, string, RawAttrCase]>(
      ([attr, held]) => (held === null ? [] : [[type, attr, held]])
    )
  );
}

const CASES = [...casesOf(NODE_FRAGMENTS), ...casesOf(MARK_FRAGMENTS)];

describe("every fragment the writer writes from is named here", () => {
  it.each(Object.entries(NODE_ATTR_ROLES))(
    "a %s says which of its attrs carry XML",
    (name, roles) => {
      expect(Object.keys(NODE_FRAGMENTS[name] ?? {}).sort()).toEqual(
        sourceAttrs(roles)
      );
    }
  );

  it.each(Object.entries(MARK_ATTR_ROLES))(
    "the %s mark says which of its attrs carry XML",
    (name, roles) => {
      expect(Object.keys(MARK_FRAGMENTS[name] ?? {}).sort()).toEqual(
        sourceAttrs(roles)
      );
    }
  );

  it("names no type the schema does not have", () => {
    expect(Object.keys(NODE_FRAGMENTS).sort()).toEqual(
      Object.keys(NODE_ATTR_ROLES).sort()
    );
    expect(Object.keys(MARK_FRAGMENTS).sort()).toEqual(
      Object.keys(MARK_ATTR_ROLES).sort()
    );
  });
});

describe("a fragment drawn into an attr and read back", () => {
  it.each(CASES)("a %s draws its %s as the attr the rule reads", (...args) => {
    const held = args[2];

    expect(
      render(held.draw(held.sound)).querySelector(`[${held.attribute}]`)
    ).not.toBeNull();
  });

  it.each(CASES)("a %s keeps a sound %s", (...args) => {
    const held = args[2];

    expect(attrStrings(roundTrip(held.draw(held.sound)))).toContain(held.sound);
  });

  it.each(CASES)("a %s turns down a %s that smuggles", (...args) => {
    const held = args[2];

    expect(attrStrings(roundTrip(held.draw(held.adversarial)))).not.toContain(
      held.adversarial
    );
  });
});

/**
 * A reply carries two fragments rather than one, and only the body of it stands in the table
 * above. The extended properties are what settles whether a thread reads as resolved, and they go
 * out into `word/commentsExtended.xml` on their own.
 */
describe("the extended properties a reply carries", () => {
  const reply = (extensionXml: string) =>
    inline(
      docxSchema.nodes.commentReference.create({
        id: "7",
        replies: [
          {
            id: "8",
            paraId: "0B0B0B0B",
            parentParaId: "0A0A0A0A",
            commentXml: '<w:comment w:id="8"><w:p/></w:comment>',
            extensionXml,
          },
        ],
      })
    );

  it("come back when they hold their shape", () => {
    const sound = '<w15:commentEx w15:paraId="0B0B0B0B" w15:done="1"/>';

    expect(attrStrings(roundTrip(reply(sound)))).toContain(sound);
  });

  it("take the whole reference down when they do not", () => {
    const smuggled = withSibling('<w15:commentEx w15:paraId="0B0B0B0B"/>');

    expect(attrStrings(roundTrip(reply(smuggled)))).not.toContain(smuggled);
  });
});
