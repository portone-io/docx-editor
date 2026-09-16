// @vitest-environment node
import { describe, expect, it } from "vitest";
import { NAMESPACES, W_NS } from "./names";
import { withEditorPrefixes } from "./partPrefixes";

const MC_NS = NAMESPACES.mc;
const W14_NS = NAMESPACES.w14;

describe("withEditorPrefixes", () => {
  it("spells a part that bound the namespace to another prefix under w", () => {
    const part =
      `<ns0:document xmlns:ns0="${W_NS}"><ns0:body>` +
      '<ns0:p><ns0:r><ns0:rPr><ns0:b ns0:val="1"/></ns0:rPr>' +
      "<ns0:t>Text</ns0:t></ns0:r></ns0:p></ns0:body></ns0:document>";
    expect(withEditorPrefixes(part)).toBe(
      `<w:document xmlns:w="${W_NS}"><w:body>` +
        '<w:p><w:r><w:rPr><w:b w:val="1"/></w:rPr>' +
        "<w:t>Text</w:t></w:r></w:p></w:body></w:document>"
    );
  });

  /**
   * The spelling of a namespace is markup, and everything that is not markup says what it says:
   * a prefix written inside a `w:t` is the document's text and a prefix inside a value is a value.
   */
  it("leaves text content and attribute values as they stand", () => {
    const part =
      `<ns0:document xmlns:ns0="${W_NS}"><ns0:body><ns0:p><ns0:r>` +
      '<ns0:t ns0:val="ns0:x">ns0:p and w:p</ns0:t>' +
      "</ns0:r></ns0:p></ns0:body></ns0:document>";
    const rewritten = withEditorPrefixes(part);
    expect(rewritten).toContain('<w:t w:val="ns0:x">ns0:p and w:p</w:t>');
  });

  it("gives the elements of a default-namespace part the prefix, and its attributes none", () => {
    const part =
      `<document xmlns="${W_NS}"><body><p><r>` +
      '<t xml:space="preserve">Text</t></r></p></body></document>';
    expect(withEditorPrefixes(part)).toBe(
      `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r>` +
        '<w:t xml:space="preserve">Text</w:t></w:r></w:p></w:body></w:document>'
    );
  });

  it("leaves a namespace it has no prefix for under the prefix the part gave it", () => {
    const part =
      `<ns0:document xmlns:ns0="${W_NS}" xmlns:o="urn:schemas-microsoft-com:office:office">` +
      "<ns0:body><ns0:p><o:p/></ns0:p></ns0:body></ns0:document>";
    expect(withEditorPrefixes(part)).toBe(
      `<w:document xmlns:w="${W_NS}" xmlns:o="urn:schemas-microsoft-com:office:office">` +
        "<w:body><w:p><o:p/></w:p></w:body></w:document>"
    );
  });

  /**
   * The part says two things by one prefix, and it is the rewrite that would make them collide, so
   * the part is handed back as it arrived. Nothing here turns a part down: what turns that one
   * down is the reader that needs the prefix, or the writer about to spell it.
   */
  it("leaves a part that has already bound a prefix the rewrite would spell out", () => {
    const takenByAnother =
      `<ns0:document xmlns:ns0="${W_NS}" xmlns:w="urn:junk">` +
      "<ns0:body/></ns0:document>";
    expect(withEditorPrefixes(takenByAnother)).toBe(takenByAnother);
    const takenUnderTheDefault = `<document xmlns="${W_NS}" xmlns:w="urn:junk"><body/></document>`;
    expect(withEditorPrefixes(takenUnderTheDefault)).toBe(takenUnderTheDefault);
  });

  it("leaves a part binding a prefix it never rewrites to whatever that part says", () => {
    const part = `<w:document xmlns:w="${W_NS}" xmlns:r="urn:foreign"><w:body/></w:document>`;
    expect(withEditorPrefixes(part)).toBe(part);
  });

  it("leaves a part whose element below the root binds a prefix the rewrite spells out", () => {
    const rebinding =
      `<ns0:document xmlns:ns0="${W_NS}"><ns0:body>` +
      '<ns0:p xmlns:w="urn:other"/></ns0:body></ns0:document>';
    expect(withEditorPrefixes(rebinding)).toBe(rebinding);
    const rebindingDefault = `<document xmlns="${W_NS}"><body><p xmlns="urn:other"/></body></document>`;
    expect(withEditorPrefixes(rebindingDefault)).toBe(rebindingDefault);
    const undeclaringTheDefault = `<document xmlns="${W_NS}"><body><p xmlns=""/></body></document>`;
    expect(withEditorPrefixes(undeclaringTheDefault)).toBe(
      undeclaringTheDefault
    );
  });

  /**
   * A declaration that names the namespace the prefix already stands for says nothing new, whether
   * it stands at the root or on a subtree: a serializer that writes a declaration on every element
   * it cut out, and a producer declaring a namespace where it uses it, both write these.
   */
  it("rewrites a part that redeclares below the root what its root already binds", () => {
    const redeclared =
      `<ns0:document xmlns:ns0="${W_NS}"><ns0:body>` +
      `<ns0:p xmlns:w="${W_NS}"><ns0:r xmlns:ns0="${W_NS}"/></ns0:p>` +
      "</ns0:body></ns0:document>";
    expect(withEditorPrefixes(redeclared)).toBe(
      `<w:document xmlns:w="${W_NS}"><w:body>` +
        `<w:p xmlns:w="${W_NS}"><w:r xmlns:w="${W_NS}"/></w:p>` +
        "</w:body></w:document>"
    );
    const redeclaredDefault = `<document xmlns="${W_NS}"><body><p xmlns="${W_NS}"/></body></document>`;
    expect(withEditorPrefixes(redeclaredDefault)).toBe(
      `<w:document xmlns:w="${W_NS}"><w:body><w:p xmlns:w="${W_NS}"/></w:body></w:document>`
    );
  });

  it("collapses a redeclaration below the root into the one name it would spell twice", () => {
    const part =
      `<ns0:document xmlns:ns0="${W_NS}"><ns0:body>` +
      `<ns0:p xmlns:ns0="${W_NS}" xmlns:w="${W_NS}"/></ns0:body></ns0:document>`;
    expect(withEditorPrefixes(part)).toBe(
      `<w:document xmlns:w="${W_NS}"><w:body>` +
        `<w:p xmlns:w="${W_NS}"/></w:body></w:document>`
    );
  });

  it("rewrites the prefix tokens of the compatibility attributes and nothing else", () => {
    const part =
      `<ns0:document xmlns:ns0="${W_NS}" xmlns:ns1="${W14_NS}" xmlns:ns2="${MC_NS}"` +
      ' ns2:Ignorable="ns1 wp14">' +
      '<ns0:body><ns0:p><ns0:r><ns0:t xml:space="preserve">Text</ns0:t>' +
      "</ns0:r></ns0:p></ns0:body></ns0:document>";
    const rewritten = withEditorPrefixes(part);
    expect(rewritten).toContain('mc:Ignorable="w14 wp14"');
    expect(rewritten).toContain('<w:t xml:space="preserve">Text</w:t>');
  });

  /**
   * `mc:MustUnderstand` names the prefixes a reader has to know (ECMA-376 Part 3 §9.5), so a token
   * left naming the prefix the part arrived under would name one nothing declares any more.
   */
  it("rewrites the prefixes mc:MustUnderstand names", () => {
    const part =
      `<ns0:document xmlns:ns0="${W_NS}" xmlns:ns2="${MC_NS}">` +
      '<ns0:body><ns0:p ns2:MustUnderstand="ns0"/></ns0:body></ns0:document>';
    expect(withEditorPrefixes(part)).toContain('<w:p mc:MustUnderstand="w"/>');
  });

  /**
   * Half a rewrite leaves the tail spelled under a prefix the rewritten root no longer binds, and
   * a part this editor never parses would carry those bytes into the exported package.
   */
  it("hands back a part whose markup cannot be read as it arrived", () => {
    const unreadableTag = `<ns0:root xmlns:ns0="${W_NS}"><ns0:a/>a < b<ns0:b/></ns0:root>`;
    expect(withEditorPrefixes(unreadableTag)).toBe(unreadableTag);
    const unreadableAttrs = `<ns0:root xmlns:ns0="${W_NS}"><ns0:p ns0:val=1/></ns0:root>`;
    expect(withEditorPrefixes(unreadableAttrs)).toBe(unreadableAttrs);
  });

  it("carries a comment, a CDATA section and a processing instruction through untouched", () => {
    const part =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<ns0:document xmlns:ns0="${W_NS}"><ns0:body>` +
      "<!-- <ns0:p/> --><![CDATA[<ns0:p/>]]><?word <ns0:p/>?>" +
      "<ns0:p/></ns0:body></ns0:document>";
    const rewritten = withEditorPrefixes(part);
    expect(rewritten).toContain("<!-- <ns0:p/> -->");
    expect(rewritten).toContain("<![CDATA[<ns0:p/>]]>");
    expect(rewritten).toContain("<?word <ns0:p/>?>");
    expect(rewritten).toContain("<w:p/>");
  });

  /** A root cannot carry one name twice, so the declaration it already has is the one kept */
  it("collapses a second declaration of one namespace into the prefix it writes", () => {
    const part =
      `<ns0:document xmlns:ns0="${W_NS}" xmlns:w="${W_NS}">` +
      "<ns0:body><w:p/></ns0:body></ns0:document>";
    expect(withEditorPrefixes(part)).toBe(
      `<w:document xmlns:w="${W_NS}"><w:body><w:p/></w:body></w:document>`
    );
  });

  it("hands back the same string for a part already spelled the way the writer writes", () => {
    const part = `<w:document xmlns:w="${W_NS}"><w:body><w:p/></w:body></w:document>`;
    expect(withEditorPrefixes(part)).toBe(part);
    const relationships =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Target="word/document.xml" Type="urn:x"/></Relationships>';
    expect(withEditorPrefixes(relationships)).toBe(relationships);
  });
});
