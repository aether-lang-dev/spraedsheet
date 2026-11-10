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
export const DIFFERS: Readonly<Record<string, number>> = { ACCRINT: 7, ARABIC: 8, AVERAGE: 9, CHOOSE: 1, COMBIN: 1, COMBINA: 1, CONCATENATE: 1, DATEDIF: 4, DEGREES: 4, "ERROR.TYPE": 2, EVEN: 3, EXP: 2, EXPONDIST: 1, FILTER: 1, FISHER: 3, FISHERINV: 3, GAMMALN: 2, IF: 9, ISREF: 3, MAX: 3, MIN: 3, MOD: 3, ODD: 3, POISSON: 2, QUOTIENT: 3, RADIANS: 4, RATE: 2, ROMAN: 8, SIGN: 2, SORT: 3, SUMIF: 9, T: 4, TIMEVALUE: 2, VALUE: 2, VALUETOTEXT: 2, XMATCH: 6, XNPV: 2, YEARFRAC: 4 };

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
