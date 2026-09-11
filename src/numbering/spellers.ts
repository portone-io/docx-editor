/**
 * How each number format spells the count of a list item.
 *
 * `w:numFmt` names some sixty formats (§17.18.59). The ones here are the ones a marker is drawn
 * in; a level naming any other counts in decimal instead. A format belongs to `NumberFormat`
 * because it has a speller here, so a marker can never be counted in a format nothing can write.
 */

/** One format's way of writing the count of an item */
interface NumberSpeller {
  /** Spells a count of one or more */
  spell: (count: number) => string;
  /**
   * How many counts one character of the spelling stands for, which is what bounds the length of
   * a marker: a letter format takes another character every 26 counted and a roman one another M
   * every 1000, so a crafted `w:start` of two billion would otherwise spell out millions of
   * characters. A format that writes a count digit by digit grows only as its digits do.
   */
  countsPerChar: number;
}

const ROMAN: ReadonlyArray<readonly [number, string]> = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];

function toRoman(count: number): string {
  let rest = count;
  let out = "";
  for (const [amount, sign] of ROMAN) {
    while (rest >= amount) {
      out += sign;
      rest -= amount;
    }
  }
  return out;
}

/** As in Word, the last symbol of a set is followed by the first one written twice (§17.18.59) */
function repeated(symbols: string, count: number): string {
  const index = (count - 1) % symbols.length;
  const times = Math.floor((count - 1) / symbols.length) + 1;
  return (symbols[index] ?? "").repeat(times);
}

const LATIN = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

/** The four marks the Chicago format counts in, U+002A, U+2020, U+2021 and U+00A7 (§17.18.59) */
const CHICAGO = "*†‡§";

/** 가 나 다 ... 파 하, the fourteen the Korean Ganada format counts in */
const GANADA = "가나다라마바사아자차카타파하";

/** Every digit written as a Korean numeral, zero first */
const KOREAN_DIGITS = "영일이삼사오육칠팔구";

function digits(symbols: string, count: number): string {
  return Array.from(`${count}`, (digit) => symbols[Number(digit)] ?? "").join(
    ""
  );
}

/**
 * The Chinese counting system, zero first and ten last.
 *
 * §17.18.59 names U+25CB for the zero while its own worked example writes U+3007, the ideographic
 * number zero, which is also what Word draws.
 */
const CHINESE_DIGITS = "〇一二三四五六七八九十";

/**
 * The Chinese counting system as §17.18.59 spells it out: the eleven symbols up to ten, then tens
 * built from 十, and from one hundred on the count written digit by digit.
 */
function chineseCounting(count: number): string {
  if (count <= 10) return CHINESE_DIGITS[count] ?? "";
  if (count >= 100) return digits(CHINESE_DIGITS, count);
  const tens = Math.floor(count / 10);
  const unit = count % 10;
  return (
    (tens === 1 ? "" : (CHINESE_DIGITS[tens] ?? "")) +
    CHINESE_DIGITS[10] +
    (unit === 0 ? "" : (CHINESE_DIGITS[unit] ?? ""))
  );
}

const POSITIONAL = Number.POSITIVE_INFINITY;

/** A `w:numFmt` this editor spells out. Every other one is drawn as a decimal */
export type NumberFormat =
  | "decimal"
  | "decimalZero"
  | "bullet"
  | "lowerLetter"
  | "upperLetter"
  | "lowerRoman"
  | "upperRoman"
  | "ganada"
  | "koreanDigital"
  | "chineseCounting"
  | "chicago";

/**
 * The speller of each format.
 * The table is written against the union, so a format cannot be named without one and a speller
 * cannot be written for a format the union does not hold.
 */
const NUMBER_SPELLERS: Record<NumberFormat, NumberSpeller> = {
  decimal: { spell: (count) => `${count}`, countsPerChar: POSITIONAL },
  /** The counts up to nine written with a zero in front (§17.18.59 decimalZero) */
  decimalZero: {
    spell: (count) => (count < 10 ? `0${count}` : `${count}`),
    countsPerChar: POSITIONAL,
  },
  /**
   * A bullet counts nothing: its `w:lvlText` is the symbol itself. A `%n` slot naming a bullet
   * level has no symbol to put there, so the count goes in as a decimal, as Word draws it.
   */
  bullet: { spell: (count) => `${count}`, countsPerChar: POSITIONAL },
  lowerLetter: {
    spell: (count) => repeated(LATIN, count).toLowerCase(),
    countsPerChar: LATIN.length,
  },
  upperLetter: {
    spell: (count) => repeated(LATIN, count),
    countsPerChar: LATIN.length,
  },
  lowerRoman: {
    spell: (count) => toRoman(count).toLowerCase(),
    countsPerChar: 1000,
  },
  upperRoman: { spell: toRoman, countsPerChar: 1000 },
  ganada: {
    spell: (count) => repeated(GANADA, count),
    countsPerChar: GANADA.length,
  },
  koreanDigital: {
    spell: (count) => digits(KOREAN_DIGITS, count),
    countsPerChar: POSITIONAL,
  },
  chineseCounting: { spell: chineseCounting, countsPerChar: POSITIONAL },
  chicago: {
    spell: (count) => repeated(CHICAGO, count),
    countsPerChar: CHICAGO.length,
  },
};

export function isNumberFormat(value: string | null): value is NumberFormat {
  return value !== null && Object.hasOwn(NUMBER_SPELLERS, value);
}

/**
 * The count as this format writes it, and as a decimal where spelling it out would run past the
 * length a marker is ever drawn at.
 */
export function spellNumber(
  count: number,
  format: NumberFormat,
  maxChars: number
): string {
  const speller = NUMBER_SPELLERS[format];
  if (count < 1 || count > speller.countsPerChar * maxChars) return `${count}`;
  return speller.spell(count);
}
