/**
 * Injectable display-only fallback stacks for fonts unavailable on the current platform. The
 * document's declared name remains first and exported `w:rFonts` is never changed.
 */

/**
 * One fallback stack and the font names that receive it.
 *
 * A group is a family: every name in it is drawn with the same stand-in font.
 */
export interface FontFallbackGroup {
  /** The fonts appended after the declared name, written as a CSS font-family list */
  readonly stack: string;
  /** The font names in this family. Letter case varies from document to document, so matching ignores it */
  readonly names: readonly string[];
}

/** The fallback fonts an editor draws with */
export interface FontFallbacks {
  readonly groups: readonly FontFallbackGroup[];
  /**
   * The stack for a name in no group, and the stack laid across the whole paper
   * when the document declares no font at all.
   */
  readonly defaultStack: string;
  /**
   * The font name the toolbar shows for a document that declares none.
   * It has to be the font `defaultStack` actually renders, since that is what the
   * reader sees.
   */
  readonly defaultFontName: string;
}

/**
 * The Korean families list what is available on the web (Pretendard) first, then
 * what the operating system already carries.
 * Pretendard is loaded by the consumer. This package only names it and never
 * downloads a font itself.
 */
const KOREAN_GOTHIC =
  '"Pretendard Variable", Pretendard, "Apple SD Gothic Neo", "Noto Sans KR", sans-serif';
const KOREAN_MYEONGJO = '"Noto Serif KR", AppleMyungjo, serif';

/**
 * The Japanese and Chinese families name the fonts the operating systems carry
 * first and the Noto family last.
 * No web font is assumed for them, unlike Pretendard for Korean, so a machine that
 * carries none of the names still lands on the right generic.
 */
const JAPANESE_GOTHIC =
  '"Hiragino Kaku Gothic ProN", "Yu Gothic", Meiryo, "Noto Sans JP", sans-serif';
const JAPANESE_MINCHO =
  '"Hiragino Mincho ProN", "Yu Mincho", "Noto Serif JP", serif';
const CHINESE_SIMPLIFIED_SANS =
  '"PingFang SC", "Microsoft YaHei", "Noto Sans SC", sans-serif';
const CHINESE_SIMPLIFIED_SERIF = '"Songti SC", "Noto Serif SC", serif';
const CHINESE_TRADITIONAL_SANS =
  '"PingFang TC", "Microsoft JhengHei", "Noto Sans TC", sans-serif';
const CHINESE_TRADITIONAL_SERIF = '"Songti TC", "Noto Serif TC", serif';

/**
 * The families whose documents write their text in Han characters, kana or Hangul.
 *
 * They are kept apart from the Latin ones because a font name belonging to one of
 * them decides which `w:rFonts` slots a font applies to (`docx/runProps`), not only
 * what stands in for it on screen.
 * Every family lists the name in the language of the document as well, since Word
 * records whichever of the two the machine that wrote the file was set to.
 */
const EAST_ASIAN_GROUPS: readonly FontFallbackGroup[] = [
  {
    stack: KOREAN_GOTHIC,
    names: [
      "malgun gothic",
      "맑은고딕",
      "맑은 고딕",
      "gulim",
      "gulimche",
      "굴림",
      "굴림체",
      "dotum",
      "dotumche",
      "돋움",
      "돋움체",
      "nanum gothic",
      "나눔고딕",
      "noto sans kr",
      "apple sd gothic neo",
      "함초롬돋움",
    ],
  },
  {
    stack: KOREAN_MYEONGJO,
    names: [
      "batang",
      "batangche",
      "바탕",
      "바탕체",
      "gungsuh",
      "gungsuhche",
      "궁서",
      "궁서체",
      "nanum myeongjo",
      "나눔명조",
      "명조",
      "함초롬바탕",
    ],
  },
  {
    stack: JAPANESE_GOTHIC,
    names: [
      "ms gothic",
      "ＭＳ ゴシック",
      "ms pgothic",
      "ＭＳ Ｐゴシック",
      "ms ui gothic",
      "ＭＳ ＵＩ ゴシック",
      "yu gothic",
      "yugothic",
      "游ゴシック",
      "yu gothic ui",
      "游ゴシック UI",
      "meiryo",
      "メイリオ",
      "meiryo ui",
      "hiragino kaku gothic pro",
      "hiragino kaku gothic pron",
      "hiragino sans",
      "ヒラギノ角ゴ Pro W3",
      "ヒラギノ角ゴ ProN W3",
      "ヒラギノ角ゴシック",
      "noto sans jp",
      "biz udgothic",
      "biz udpgothic",
      "BIZ UDゴシック",
    ],
  },
  {
    stack: JAPANESE_MINCHO,
    names: [
      "ms mincho",
      "ＭＳ 明朝",
      "ms pmincho",
      "ＭＳ Ｐ明朝",
      "yu mincho",
      "yumincho",
      "游明朝",
      "hiragino mincho pro",
      "hiragino mincho pron",
      "ヒラギノ明朝 Pro W3",
      "ヒラギノ明朝 ProN W3",
      "ヒラギノ明朝",
      "noto serif jp",
      "biz udmincho",
      "biz udpmincho",
      "BIZ UD明朝",
    ],
  },
  {
    stack: CHINESE_SIMPLIFIED_SANS,
    names: [
      "microsoft yahei",
      "microsoft yahei ui",
      "微软雅黑",
      "pingfang sc",
      "苹方",
      "苹方-简",
      "noto sans sc",
      "source han sans sc",
      "simhei",
      "黑体",
      "heiti sc",
      "dengxian",
      "等线",
    ],
  },
  {
    stack: CHINESE_SIMPLIFIED_SERIF,
    names: [
      "simsun",
      "宋体",
      "nsimsun",
      "新宋体",
      "songti sc",
      "fangsong",
      "仿宋",
      "kaiti",
      "楷体",
      "noto serif sc",
      "source han serif sc",
    ],
  },
  {
    stack: CHINESE_TRADITIONAL_SANS,
    names: [
      "microsoft jhenghei",
      "microsoft jhenghei ui",
      "微軟正黑體",
      "pingfang tc",
      "蘋方-繁",
      "noto sans tc",
    ],
  },
  {
    stack: CHINESE_TRADITIONAL_SERIF,
    names: [
      "pmingliu",
      "新細明體",
      "mingliu",
      "細明體",
      "dfkai-sb",
      "標楷體",
      "songti tc",
      "noto serif tc",
    ],
  },
];

/**
 * The Latin families.
 * Every one of these names has a close equivalent in every environment, so one
 * standard generic is enough for them.
 */
const LATIN_GROUPS: readonly FontFallbackGroup[] = [
  {
    stack: "sans-serif",
    names: [
      "arial",
      "helvetica",
      "verdana",
      "tahoma",
      "trebuchet ms",
      "calibri",
      "segoe ui",
      "roboto",
      "open sans",
      "noto sans",
    ],
  },
  {
    stack: "serif",
    names: [
      "times new roman",
      "georgia",
      "garamond",
      "cambria",
      "book antiqua",
      "palatino linotype",
    ],
  },
  { stack: "monospace", names: ["courier new", "consolas"] },
];

/**
 * The built-in set: the CJK and Latin families above, and a Latin default.
 *
 * An unrecognized name, and a document that declares no font at all, land on a
 * Latin sans. A screen serving documents in a particular language hands in its
 * own set (`fontFallbacks` on `DocxEditor`).
 */
export const DEFAULT_FONT_FALLBACKS: FontFallbacks = {
  groups: [...EAST_ASIAN_GROUPS, ...LATIN_GROUPS],
  defaultStack: "Arial, Helvetica, sans-serif",
  defaultFontName: "Arial",
};

/**
 * The form a name is compared in. Quotes, letter case, and Unicode normalization all vary
 * from document to document, so two spellings of one name have to compare equal.
 *
 * NFC only. An NFK* form is lossy and would fold the fullwidth Latin of a name such as
 * `ＤＦＫａｉＳｈｕ` into ASCII. The result is a comparison key and is never
 * written into a document or shown: the original spelling is what gets stored and exported.
 */
export function comparableFontName(name: string): string {
  return name
    .trim()
    .replace(/^["']|["']$/g, "")
    .toLowerCase()
    .normalize("NFC");
}

/**
 * Han characters, kana, Hangul, and the fullwidth Latin letters a name such as
 * `ＭＳ 明朝` is written with.
 * A name holding one of these is East Asian whether or not we have ever heard of it.
 */
const EAST_ASIAN_CHARACTER =
  /[\u1100-\u11ff\u3000-\u30ff\u3130-\u318f\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff\ua960-\ua97f\uac00-\ud7ff\uf900-\ufaff\uff00-\uffef\u{20000}-\u{2fa1f}]/u;

const eastAsianNames: ReadonlySet<string> = new Set(
  EAST_ASIAN_GROUPS.flatMap((group) => group.names.map(comparableFontName))
);

/**
 * Whether this font name belongs to a script written in Han characters, kana or
 * Hangul.
 *
 * It decides which `w:rFonts` slots a font applies to, so it reads the built-in
 * families rather than the fallback set the consumer handed in: which fonts stand in
 * for which names on screen must not change what is written into the document.
 */
export function isEastAsianFontName(name: string): boolean {
  return (
    EAST_ASIAN_CHARACTER.test(name) ||
    eastAsianNames.has(comparableFontName(name))
  );
}

/**
 * The stack per name, built once per fallback set.
 * A run is drawn one at a time, so the groups are not walked again for every run.
 */
const stacksByFallbacks = new WeakMap<FontFallbacks, Map<string, string>>();

function stackByName(fallbacks: FontFallbacks): Map<string, string> {
  const built = stacksByFallbacks.get(fallbacks);
  if (built) return built;

  const stacks = new Map<string, string>();
  for (const group of fallbacks.groups) {
    for (const name of group.names)
      stacks.set(comparableFontName(name), group.stack);
  }
  stacksByFallbacks.set(fallbacks, stacks);
  return stacks;
}

/** The stacks of every name the list recognizes, in the order the names appear */
function matchedStacks(
  cssNames: string,
  fallbacks: FontFallbacks
): readonly string[] {
  const stacks = stackByName(fallbacks);
  const matched: string[] = [];
  for (const name of cssNames.split(",")) {
    const stack = stacks.get(comparableFontName(name));
    if (stack !== undefined && !matched.includes(stack)) matched.push(stack);
  }
  return matched;
}

/**
 * The CSS keywords that stand for a font the reader is certain to have.
 * One of them matches every character it is asked for, so a name behind it is never
 * reached and they are all held back to the end of the list.
 */
const GENERIC_FAMILIES: ReadonlySet<string> = new Set([
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
  "ui-serif",
  "ui-sans-serif",
  "ui-monospace",
  "ui-rounded",
  "math",
  "emoji",
]);

function isGenericFamily(name: string): boolean {
  return GENERIC_FAMILIES.has(comparableFontName(name));
}

/** Naming one font twice draws no differently than naming it once, so the repeat goes */
function distinctNames(names: readonly string[]): readonly string[] {
  const seen = new Set<string>();
  return names.filter((name) => {
    const comparable = comparableFontName(name);
    if (seen.has(comparable)) return false;
    seen.add(comparable);
    return true;
  });
}

/**
 * Appends a fallback font of the same family after each name the document declares.
 *
 * A run of a Japanese or Chinese document names a Latin font and an East Asian one
 * both, one per `w:rFonts` slot, and the two belong to different families: each of
 * them needs the stack of its own family behind it.
 */
export function withFontFallback(
  cssNames: string,
  fallbacks: FontFallbacks = DEFAULT_FONT_FALLBACKS
): string {
  const matched = matchedStacks(cssNames, fallbacks);
  const appended = (
    matched.length > 0 ? matched : [fallbacks.defaultStack]
  ).flatMap((stack) => stack.split(","));
  return distinctNames([
    ...cssNames.split(","),
    ...appended.filter((name) => !isGenericFamily(name)),
    ...appended.filter(isGenericFamily),
  ]).join(",");
}
