// @vitest-environment node
import { describe, expect, it } from "vitest";
import {
  comparableFontName,
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

/**
 * A name a document declares is an untrusted string and may be normalized differently than the
 * built-in names it is compared against.
 * The forms are written as escapes so that saving this file in one normal form cannot quietly
 * turn these tests into comparisons of a name with itself.
 */
describe("a name written in a different Unicode normal form", () => {
  // "맑은 고딕", composed and decomposed
  const MALGUN_NFC = "\ub9d1\uc740 \uace0\ub515";
  const MALGUN_NFD =
    "\u1106\u1161\u11b0\u110b\u1173\u11ab \u1100\u1169\u1103\u1175\u11a8";
  // "Málaga", composed and decomposed
  const MALAGA_NFC = "M\u00e1laga";
  const MALAGA_NFD = "Ma\u0301laga";

  it("is the same name to compare against", () => {
    expect(MALGUN_NFD).not.toBe(MALGUN_NFC);
    expect(MALAGA_NFD).not.toBe(MALAGA_NFC);
    expect(comparableFontName(MALGUN_NFD)).toBe(comparableFontName(MALGUN_NFC));
    expect(comparableFontName(MALAGA_NFD)).toBe(comparableFontName(MALAGA_NFC));
  });

  it("still reaches the family the composed name belongs to", () => {
    expect(withFontFallback(`"${MALGUN_NFD}"`)).toBe(
      `"${MALGUN_NFD}",${KOREAN_GOTHIC_STACK}`
    );
  });

  it("is drawn under the spelling the document wrote, not a rewritten one", () => {
    const stack = withFontFallback(`"${MALGUN_NFD}"`);
    expect(lead(stack, `"${MALGUN_NFD}"`)).toBe(`"${MALGUN_NFD}"`);
    expect(stack).not.toContain(MALGUN_NFC);
  });

  it("is folded no further than NFC", () => {
    // NFKC would fold these onto their ASCII and halfwidth counterparts and lose the
    // distinction between two genuinely different names
    expect(
      comparableFontName("\uff24\uff26\uff2b\uff41\uff49\uff33\uff48\uff55")
    ).not.toBe(comparableFontName("DFKaiShu"));
    expect(comparableFontName("\u3231")).not.toBe(
      comparableFontName("(\u682a)")
    );
  });
});

/**
 * `isEastAsianFontName` picks the `w:rFonts` slots a font is written into, so its answer is
 * part of what a round trip exports. It must not depend on the normal form a name arrives in.
 */
describe("the slot selection a name drives", () => {
  const NAMES = [
    "\ub9d1\uc740 \uace0\ub515",
    "\ubc14\ud0d5\uccb4",
    "\uff2d\uff33 \u660e\u671d",
    "\u6e38\u30b4\u30b7\u30c3\u30af",
    "\u5fae\u8f6f\u96c5\u9ed1",
    "Malgun Gothic",
    "DFKai-SB",
    "M\u00e1laga",
    "Arial",
  ];

  it("is the same whichever normal form the document wrote the name in", () => {
    for (const name of NAMES) {
      expect(isEastAsianFontName(name.normalize("NFD"))).toBe(
        isEastAsianFontName(name.normalize("NFC"))
      );
    }
  });

  it("still reads a fullwidth Latin name as East Asian", () => {
    // NFKC would turn this into ASCII and write it into the Latin slots instead
    const fullwidth = "\uff24\uff26\uff2b\uff41\uff49\uff33\uff48\uff55";
    expect(isEastAsianFontName(fullwidth)).toBe(true);
    expect(isEastAsianFontName(fullwidth.normalize("NFD"))).toBe(true);
  });
});
