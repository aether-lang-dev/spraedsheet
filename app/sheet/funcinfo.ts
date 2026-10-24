// What the app knows of each function, in one record (docs/comparisons/HYPERFORMULA.md
// action 4, GNUMERIC.md action 1): the library's arity, the evaluator's own
// forms, volatility, the TRUE/FALSE refusal, the Spanish name, and where it is
// known to differ from Excel. formulas.functionInfo and formulas.functionStatus
// answer from here; tests/headless/sheet_funcinfo.ts checks the tables it reads
// agree with one another.
import { arityOf, names, NO_LOGICALS_LIST } from "./functions.ts";
import { SPECIAL_FORMS, VOLATILE_FORMS } from "./evaluate.ts";
import { ES_FUNCTIONS } from "./locale.ts";

export interface FunctionInfo {
  readonly name: string;
  // "function" (the library's), "form" (the evaluator's own: IF, ROW, OFFSET ...),
  // or "special form" (LET, LAMBDA, MAP ...: arguments it does not evaluate first).
  readonly kind: "function" | "form" | "special form";
  readonly min: number | null;
  readonly max: number | null;          // null: no most
  readonly arity: string;
  readonly volatile: boolean;
  readonly refusesLogicals: boolean;    // TRUE and FALSE are #VALUE!, as Excel gives
  readonly es: string | null;
  // Where it is known to differ from Excel, or null.
  readonly deviation: string | null;
}

// Cells of IronCalc's Excel-authored workbooks (third_party/ironcalc) that this
// app still works out differently from the value Excel stored, by function,
// each workbook counted for the functions its name gives (tests/headless/
// lib/ironcalc_calc_known.ts is the ledger; sheet_funcinfo recounts it).
export const DIFFERS: Readonly<Record<string, number>> = { ACCRINT: 9, ACCRINTM: 1, AND: 12, ARABIC: 8, ARRAYTOTEXT: 1, AVERAGE: 17, CHOOSE: 3, COMBIN: 1, COMBINA: 1, CONCATENATE: 1, CONFIDENCE: 2, COUNT: 9, DATE: 5, DATEDIF: 12, DATEVALUE: 5, DAY: 5, DB: 5, DDB: 5, DEGREES: 4, DROP: 5, EDATE: 6, EOMONTH: 7, "ERROR.TYPE": 6, EVEN: 3, EXP: 12, EXPONDIST: 1, FILTER: 1, FIND: 3, FISHER: 3, FISHERINV: 3, FORMULATEXT: 6, FREQUENCY: 1, FV: 1, GAMMALN: 2, GROWTH: 6, IF: 10, IFS: 5, INDEX: 6, IPMT: 4, ISEVEN: 1, ISFORMULA: 1, ISODD: 1, ISOMITTED: 2, ISOWEEKNUM: 8, ISREF: 11, LAMBDA: 5, LARGE: 3, LINEST: 4, LN: 4, LOG: 4, LOG10: 4, MATCH: 9, MAX: 4, MAXA: 3, MIN: 4, MINA: 3, MOD: 5, MONTH: 5, MULTINOMIAL: 4, N: 1, ODD: 3, OR: 12, PMT: 9, POISSON: 3, POWER: 5, PPMT: 4, PROB: 1, PRODUCT: 6, QUOTIENT: 5, RADIANS: 4, RATE: 2, ROMAN: 8, SEARCH: 3, SEQUENCE: 4, SERIESSUM: 2, SIGN: 12, SMALL: 3, SORT: 3, SUM: 6, SUMIF: 9, SUMSQ: 1, T: 14, TAKE: 5, TIMEVALUE: 6, TREND: 6, VALUE: 12, VALUETOTEXT: 12, WEEKDAY: 8, WEEKNUM: 8, XMATCH: 6, XNPV: 7, XOR: 12, YEAR: 5, YEARFRAC: 12 };

const plural = (n: number) => `${n} value${n === 1 ? "" : "s"}`;
const takes = (min: number, max: number | null): string =>
  max === null ? `at least ${plural(min)}` : max === min ? (min === 0 ? "no values" : plural(min)) : max === min + 1 ? `${min} or ${plural(max)}` : `${min} to ${plural(max)}`;

export function functionInfo(name: string): FunctionInfo | null {
  const n = name.trim().toUpperCase();
  if (!names().includes(n) && !SPECIAL_FORMS.includes(n)) return null;
  const a = arityOf(n);
  const special = SPECIAL_FORMS.includes(n);
  const cells = DIFFERS[n];
  return {
    name: n, kind: special ? "special form" : a ? "function" : "form",
    min: a ? a.min : null, max: a && a.max !== Infinity ? a.max : null, arity: a ? takes(a.min, a.max === Infinity ? null : a.max) : "",
    volatile: VOLATILE_FORMS.has(n), refusesLogicals: NO_LOGICALS_LIST.includes(n), es: ES_FUNCTIONS[n] ?? null,
    deviation: cells ? `${cells} cell${cells === 1 ? "" : "s"} of IronCalc's Excel-authored workbooks differ from Excel's stored value (docs/comparisons/COMPATIBILITY.md)` : null,
  };
}

export function allFunctions(): FunctionInfo[] {
  return [...new Set([...names(), ...SPECIAL_FORMS])].sort().map((n) => functionInfo(n)!);
}
