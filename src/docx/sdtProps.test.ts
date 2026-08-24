// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { parseXml, W_NS } from "../ooxml/xml";
import { readSdtWrapper } from "./sdt";
import { withContentLock } from "./sdtProps";

const ALIAS = '<w:alias w:val="signedOn"/>';
const TAG = '<w:tag w:val="date"/>';
const ID = '<w:id w:val="7"/>';
const BINDING = '<w:dataBinding w:xpath="/root/date"/>';
const LOCK = '<w:lock w:val="sdtContentLocked"/>';

const prefix = (props: string) => `<w:sdt><w:sdtPr>${props}</w:sdtPr>`;

/** The control the prefix opens, closed up so it can be read back in */
function control(sdtPrefix: string): Element {
  const xml = `${sdtPrefix}<w:sdtContent><w:p/></w:sdtContent></w:sdt>`;
  const parsed = parseXml(`<w:wrap xmlns:w="${W_NS}">${xml}</w:wrap>`);
  const el = parsed.documentElement.firstElementChild;
  if (!el) throw new Error("no element");
  return el;
}

describe("shutting a control", () => {
  it("writes the lock into the spot the order calls for", () => {
    expect(withContentLock(prefix(ALIAS + TAG + ID + BINDING), true)).toBe(
      prefix(ALIAS + TAG + ID + LOCK + BINDING)
    );
  });

  it("replaces a lock that only kept the control from being deleted", () => {
    expect(
      withContentLock(prefix(`${ID}<w:lock w:val="sdtLocked"/>`), true)
    ).toBe(prefix(ID + LOCK));
  });

  it("comes back out as a control that says its contents are shut", () => {
    const shut = withContentLock(prefix(ID), true);
    if (shut === null) throw new Error("the control could not be rewritten");
    expect(readSdtWrapper(control(shut))?.contentsLocked).toBe(true);
  });
});

describe("lifting a control's lock", () => {
  it("leaves everything else the control carries exactly as it was", () => {
    const locked =
      '<w:sdt w:rsidR="00A">' +
      `<w:sdtPr>${ALIAS}${TAG}${ID}${LOCK}${BINDING}</w:sdtPr>` +
      "<w:sdtEndPr><w:rPr/></w:sdtEndPr>";
    expect(withContentLock(locked, false)).toBe(
      '<w:sdt w:rsidR="00A">' +
        `<w:sdtPr>${ALIAS}${TAG}${ID}${BINDING}</w:sdtPr>` +
        "<w:sdtEndPr><w:rPr/></w:sdtEndPr>"
    );
  });

  /** A control with no `w:sdtPr` at all is one we no longer read back, so an emptied one stays written */
  it("keeps the properties element of a control that carried nothing but the lock", () => {
    const opened = withContentLock(prefix(LOCK), false);
    expect(opened).toBe("<w:sdt><w:sdtPr/>");
    if (opened === null) throw new Error("the control could not be rewritten");
    expect(readSdtWrapper(control(opened))?.contentsLocked).toBe(false);
  });
});

describe("a control opening we cannot make out", () => {
  it("is handed back as null, so the caller can leave it alone", () => {
    expect(withContentLock("<w:sdt>", true)).toBeNull();
    expect(withContentLock("<w:sdt><w:sdtPr>", false)).toBeNull();
    expect(withContentLock("lock", true)).toBeNull();
  });
});
