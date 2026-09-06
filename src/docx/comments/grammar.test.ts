// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { parseXml, W_NS } from "../../ooxml/xml";
import { W14_NS } from "./constants";
import {
  readStrictCommentBody,
  renderCommentBody,
  withThreadKey,
} from "./grammar";

const entry = (body: string): Element =>
  parseXml(
    `<w:comment xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" w:id="0">${body}</w:comment>`
  ).documentElement;

/**
 * The writer and the reader are two halves of one grammar, and a file is judged by whether the
 * half that reads recognises what the half that writes put out. Holding them against each other is
 * what keeps them from drifting apart.
 */
describe("what the writer writes, the reader reads back", () => {
  it.each([
    ["plain", "note"],
    ["empty", ""],
    ["a line break", "first\nsecond"],
    ["a blank line", "first\n\nsecond"],
    ["a leading break", "\nsecond"],
    ["a trailing break", "first\n"],
    ["kept space", "  padded  "],
    ["markup characters", 'x < y & z > "w"'],
    ["not latin", "댓글 \u{1F600}"],
  ])("reads %s back as it was written", (_what, text) => {
    for (const key of [null, "12345678"]) {
      expect(readStrictCommentBody(entry(renderCommentBody(text, key)))).toBe(
        text
      );
    }
  });

  it("writes a body the reader takes and nothing it does not", () => {
    expect(renderCommentBody("note", null)).toBe(
      '<w:p><w:r><w:t xml:space="preserve">note</w:t></w:r></w:p>'
    );
    expect(renderCommentBody("a\nb", "12345678")).toBe(
      `<w:p xmlns:w14="${W14_NS}" w14:paraId="12345678">` +
        '<w:r><w:t xml:space="preserve">a</w:t><w:br/>' +
        '<w:t xml:space="preserve">b</w:t></w:r></w:p>'
    );
  });
});

describe("giving an entry a thread key", () => {
  const body = '<w:p><w:r><w:t xml:space="preserve">note</w:t></w:r></w:p>';

  it("puts the key on the body and leaves the rest of the entry alone", () => {
    expect(
      withThreadKey(`<w:comment w:id="0">${body}</w:comment>`, "ABCD1234")
    ).toBe(
      '<w:comment w:id="0"><w:p w14:paraId="ABCD1234">' +
        '<w:r><w:t xml:space="preserve">note</w:t></w:r></w:p></w:comment>'
    );
  });

  it("leaves an entry that already carries one exactly as it was", () => {
    const already = `<w:comment w:id="0"><w:p w14:paraId="0000AAAA">${body.slice(5)}</w:comment>`;
    expect(withThreadKey(already, "ABCD1234")).toBe(already);
  });

  it("puts it on the last paragraph, which is the one the key is read off", () => {
    const two = `<w:comment w:id="0">${body}${body}</w:comment>`;
    const keyed = withThreadKey(two, "ABCD1234");

    expect(keyed.indexOf("w14:paraId")).toBeGreaterThan(
      keyed.indexOf("</w:p>")
    );
    expect(keyed.match(/w14:paraId/g)).toHaveLength(1);
  });

  it("keeps a body the writer would not have written, rather than putting one back", () => {
    const rich =
      '<w:comment w:id="0"><w:p><w:r><w:rPr><w:b/></w:rPr>' +
      '<w:t xml:space="preserve">bold</w:t></w:r></w:p></w:comment>';

    expect(withThreadKey(rich, "ABCD1234")).toContain("<w:b/>");
  });
});
