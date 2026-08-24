// @vitest-environment node
import { describe, expect, it } from "vitest";
import {
  DEFAULT_FONT_FALLBACKS,
  type FontFallbacks,
  isEastAsianFontName,
  withFontFallback,
} from "./fontStack";

const DEFAULT_FONT_STACK = DEFAULT_FONT_FALLBACKS.defaultStack;
const KOREAN_GOTHIC_STACK =
  '"Pretendard Variable", Pretendard, "Apple SD Gothic Neo", "Noto Sans KR", sans-serif';

/** The leading part with the fallbacks removed. The name the document declared has to survive untouched */
function lead(stack: string, names: string): string {
  return stack.startsWith(`${names},`) ? names : stack;
}

describe("withFontFallback", () => {
  it("leaves the name the document wrote down at the very front untouched", () => {
    const stack = withFontFallback('"Malgun Gothic"');
    expect(lead(stack, '"Malgun Gothic"')).toBe('"Malgun Gothic"');
  });

  it("a Korean gothic name is followed by the gothic family", () => {
    expect(withFontFallback('"Malgun Gothic"')).toContain("Pretendard");
    expect(withFontFallback('"Malgun Gothic"')).toContain(
      '"Apple SD Gothic Neo"'
    );
    expect(withFontFallback('"굴림"')).toContain('"Noto Sans KR"');
    expect(withFontFallback('"Dotum"')).toBe(`"Dotum",${KOREAN_GOTHIC_STACK}`);
  });

  it("a Korean myungjo name is followed by the myungjo family", () => {
    const stack = withFontFallback('"Batang"');
    expect(stack).toContain('"Noto Serif KR"');
    expect(stack.endsWith("serif")).toBe(true);
    expect(stack).not.toContain("sans-serif");
    expect(withFontFallback('"바탕체"')).toContain("AppleMyungjo");
  });

  it("appends only a single standard generic to a Latin font", () => {
    expect(withFontFallback('"Arial"')).toBe('"Arial",sans-serif');
    expect(withFontFallback('"Georgia"')).toBe('"Georgia",serif');
    expect(withFontFallback('"Courier New"')).toBe('"Courier New",monospace');
  });

  it("reads a name as the same family regardless of letter case", () => {
    expect(withFontFallback('"MALGUN GOTHIC"')).toBe(
      `"MALGUN GOTHIC",${KOREAN_GOTHIC_STACK}`
    );
  });

  it("every name of a list is followed by the family it belongs to", () => {
    const stack = withFontFallback('"Batang","Arial"');
    expect(stack.startsWith('"Batang","Arial",')).toBe(true);
    expect(stack).toContain('"Noto Serif KR"');
    expect(stack).toContain("sans-serif");
  });

  it("backs a Latin name and a mincho name each with its own family", () => {
    const stack = withFontFallback('"Times New Roman","MS Mincho"');
    expect(stack.startsWith('"Times New Roman","MS Mincho",')).toBe(true);
    expect(stack).toContain('"Hiragino Mincho ProN"');
    expect(stack).toContain('"Yu Mincho"');
    expect(stack).toContain('"Noto Serif JP"');
    // Left to the first name alone, the mincho families were never named and the
    // Han characters were drawn with whatever the Latin generic stands for
    expect(stack).toBe(
      '"Times New Roman","MS Mincho","Hiragino Mincho ProN", "Yu Mincho", "Noto Serif JP",serif'
    );
  });

  it("holds the generic keywords back behind the names of real fonts", () => {
    const stack = withFontFallback('"Arial","SimSun"');
    expect(stack.indexOf('"Songti SC"')).toBeLessThan(
      stack.indexOf("sans-serif")
    );
    expect(stack).toBe(
      '"Arial","SimSun","Songti SC", "Noto Serif SC",sans-serif, serif'
    );
  });

  it("names a font once even where the list and the family it pulls in share it", () => {
    expect(withFontFallback('"Noto Sans KR"')).toBe(
      '"Noto Sans KR","Pretendard Variable", Pretendard, "Apple SD Gothic Neo", sans-serif'
    );
  });

  it.each(['"Unlisted Display"', '"Unlisted Display","Another Unlisted"'])(
    "appends the default stack where no name is one we know: %s",
    (names) => {
      expect(withFontFallback(names)).toBe(`${names},${DEFAULT_FONT_STACK}`);
    }
  );

  it("reads past a name it does not know to the one it does", () => {
    const stack = withFontFallback('"Unlisted Display","MS Mincho"');
    expect(stack).toContain('"Hiragino Mincho ProN"');
    expect(stack).not.toContain("Pretendard");
  });

  it("a Japanese gothic name is followed by the Japanese gothic family", () => {
    expect(withFontFallback('"MS Gothic"')).toContain(
      '"Hiragino Kaku Gothic ProN"'
    );
    expect(withFontFallback('"MS PGothic"')).toContain('"Yu Gothic"');
    expect(withFontFallback('"ＭＳ ゴシック"')).toContain("Meiryo");
    expect(withFontFallback('"メイリオ"')).toContain('"Noto Sans JP"');
    expect(withFontFallback('"BIZ UDGothic"').endsWith("sans-serif")).toBe(
      true
    );
  });

  it("a Japanese mincho name is followed by the Japanese mincho family", () => {
    const stack = withFontFallback('"MS Mincho"');
    expect(stack).toContain('"Hiragino Mincho ProN"');
    expect(stack).toContain('"Noto Serif JP"');
    expect(stack.endsWith("serif")).toBe(true);
    expect(stack).not.toContain("sans-serif");
    // Left unrecognized it used to fall through to the Korean gothic stack, which
    // drew a mincho document in a sans-serif font
    expect(stack).not.toContain("Pretendard");
    expect(withFontFallback('"ＭＳ 明朝"')).toContain('"Yu Mincho"');
    expect(withFontFallback('"游明朝"')).toContain('"Noto Serif JP"');
  });

  it("a simplified Chinese name is followed by the family of its own script", () => {
    expect(withFontFallback('"Microsoft YaHei"')).toContain('"PingFang SC"');
    expect(withFontFallback('"微软雅黑"')).toContain('"Noto Sans SC"');
    expect(withFontFallback('"SimHei"').endsWith("sans-serif")).toBe(true);

    const serif = withFontFallback('"SimSun"');
    expect(serif).toContain('"Songti SC"');
    expect(serif.endsWith("serif")).toBe(true);
    expect(serif).not.toContain("sans-serif");
    expect(withFontFallback('"宋体"')).toContain('"Noto Serif SC"');
    expect(withFontFallback('"楷体"')).toContain('"Songti SC"');
  });

  it("a traditional Chinese name is followed by the traditional family", () => {
    expect(withFontFallback('"Microsoft JhengHei"')).toContain('"PingFang TC"');
    expect(withFontFallback('"微軟正黑體"')).toContain('"Noto Sans TC"');
    expect(withFontFallback('"PMingLiU"')).toContain('"Songti TC"');
    expect(withFontFallback('"標楷體"')).toContain('"Noto Serif TC"');
    expect(withFontFallback('"MingLiU"').endsWith("serif")).toBe(true);
  });

  it("reads a CJK name as the same family regardless of letter case", () => {
    const mincho =
      '"Hiragino Mincho ProN", "Yu Mincho", "Noto Serif JP", serif';
    expect(withFontFallback('"MS MINCHO"')).toBe(`"MS MINCHO",${mincho}`);
    expect(withFontFallback('"ms mincho"')).toBe(`"ms mincho",${mincho}`);
    expect(withFontFallback('"SIMSUN"')).toContain('"Songti SC"');
  });
});

/**
 * A consumer that loads its own web fonts hands in its own set.
 * Nothing about the rule changes: the declared name stays at the front and the
 * stack of the family it belongs to follows.
 */
describe("fallback fonts handed in by the consumer", () => {
  const DEVANAGARI: FontFallbacks = {
    groups: [
      {
        stack: '"Noto Sans Devanagari", sans-serif',
        names: ["Mangal", "Kokila"],
      },
      { stack: '"Noto Serif Devanagari", serif', names: ["Sanskrit Text"] },
    ],
    defaultStack: '"Noto Sans Devanagari", sans-serif',
    defaultFontName: "Noto Sans Devanagari",
  };

  it("uses the map that was handed in as is", () => {
    expect(withFontFallback('"Mangal"', DEVANAGARI)).toBe(
      '"Mangal","Noto Sans Devanagari", sans-serif'
    );
    expect(withFontFallback('"Sanskrit Text"', DEVANAGARI)).toBe(
      '"Sanskrit Text","Noto Serif Devanagari", serif'
    );
  });

  it("matches a name even when it is written in upper case", () => {
    expect(withFontFallback('"KOKILA"', DEVANAGARI)).toBe(
      '"KOKILA","Noto Sans Devanagari", sans-serif'
    );
  });

  it("no longer recognizes the names from the default map", () => {
    expect(withFontFallback('"Batang"', DEVANAGARI)).toBe(
      '"Batang","Noto Sans Devanagari", sans-serif'
    );
  });

  it("the map handed in does not touch the default map", () => {
    expect(withFontFallback('"Batang"')).toContain('"Noto Serif KR"');
  });
});

describe("isEastAsianFontName", () => {
  it("recognizes the names of the built-in East Asian families", () => {
    expect(isEastAsianFontName("Malgun Gothic")).toBe(true);
    expect(isEastAsianFontName("Batang")).toBe(true);
    expect(isEastAsianFontName("MS Mincho")).toBe(true);
    expect(isEastAsianFontName("Yu Gothic UI")).toBe(true);
    expect(isEastAsianFontName("SimSun")).toBe(true);
    expect(isEastAsianFontName("Microsoft JhengHei")).toBe(true);
    expect(isEastAsianFontName("DFKai-SB")).toBe(true);
  });

  it("ignores letter case and the quotes a CSS name list carries", () => {
    expect(isEastAsianFontName("ms gothic")).toBe(true);
    expect(isEastAsianFontName("MS GOTHIC")).toBe(true);
    expect(isEastAsianFontName('"PMingLiU"')).toBe(true);
  });

  it("recognizes a name written in the script itself, listed or not", () => {
    expect(isEastAsianFontName("ＭＳ 明朝")).toBe(true);
    expect(isEastAsianFontName("游ゴシック")).toBe(true);
    expect(isEastAsianFontName("微软雅黑")).toBe(true);
    // A name nobody has listed still writes its text in Han characters
    expect(isEastAsianFontName("架空明朝体")).toBe(true);
    expect(isEastAsianFontName("カナ見本")).toBe(true);
  });

  it("a Latin name is not East Asian", () => {
    expect(isEastAsianFontName("Arial")).toBe(false);
    expect(isEastAsianFontName("Times New Roman")).toBe(false);
    expect(isEastAsianFontName("Courier New")).toBe(false);
    // An unlisted Latin name too, since nothing about it says otherwise
    expect(isEastAsianFontName("Unlisted Display")).toBe(false);
  });
});
