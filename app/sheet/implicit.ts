// Implicit intersection as Excel's loader puts it. A formula a file stores with
// no array flag was worked out by an Excel that intersected implicitly: a range
// where one value goes gave the cell of it in the formula's own row or column.
// The Excel that opens it today shows that with an @ (=@A:A, =ROMAN(@H34:H37),
// =@INDEX(...)) and works it out to one value; this app spills a formula, so
// the importer puts the @ in (intersected), and the writer takes the ones it
// would put in out again (stored), so a formula comes back as the legacy one it was.
// An array formula (a file's t="array", which is also how it keeps a dynamic one)
// is left as it is.
//
// Where the @ goes: before a range (A1:A3, A:A, a name of a range, a table's column)
// where an operand or a function's argument of one value is, and before a call that
// may give a range or an array (INDEX, OFFSET, INDIRECT, TRANSPOSE ...) where one
// value is. Not before what is a value already, nor where the function takes a
// range itself (SUM(A1:A3)), nor in what an array-class argument works out
// (SUMPRODUCT(A1:A3*B1:B3): that is worked out as arrays, as Excel does).
import { tokens } from "./lexer.ts";
import type { Token } from "./lexer.ts";
import { takesOneAt } from "./functions.ts";

type Node =
  | { readonly t: "other" }
  | { readonly t: "range"; readonly start: number }
  | { readonly t: "call"; readonly start: number; readonly name: string; readonly args: readonly Node[] }
  | { readonly t: "bin"; readonly left: Node; readonly right: Node }
  | { readonly t: "un"; readonly arg: Node }
  | { readonly t: "paren"; readonly inner: Node };

// Calls that may give a range or an array: Excel puts the @ before one where one value goes.
const MAY_GIVE_RANGE = new Set(["INDEX", "OFFSET", "INDIRECT", "TRANSPOSE", "MMULT", "MINVERSE", "MUNIT", "LINEST", "LOGEST", "TREND", "GROWTH", "FREQUENCY",
  "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "TEXTSPLIT", "VSTACK", "HSTACK", "TOCOL", "TOROW", "TAKE", "DROP", "EXPAND", "WRAPROWS", "WRAPCOLS", "CHOOSECOLS", "CHOOSEROWS"]);

// What a function's arguments are to the loader: V (one value), R (a range
// or a value: a bare range is left, an operation in it is of one value each), A
// (an array, worked out as such), or the call's own class (IF's branches).
type Class = "V" | "R" | "A" | "self";
const ARRAY_CLASS = new Set(["SUMPRODUCT", "MMULT", "TRANSPOSE", "MINVERSE", "MDETERM", "LINEST", "LOGEST", "TREND", "GROWTH", "FREQUENCY", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "LOOKUP",
  "GROUPBY", "PIVOTBY", "FILTER", "SORT", "SORTBY", "UNIQUE", "TEXTSPLIT", "VSTACK", "HSTACK", "TOCOL", "TOROW", "TAKE", "DROP", "EXPAND", "WRAPROWS", "WRAPCOLS", "CHOOSECOLS", "CHOOSEROWS"]);
const RANGE_ARGS = new Set(["ROW", "COLUMN", "ROWS", "COLUMNS", "AREAS", "ISREF", "SHEET", "SHEETS", "FORMULATEXT", "ISFORMULA", "CELL", "ISOMITTED", "INFO"]);

function classOf(name: string, nargs: number, i: number): Class {
  if (ARRAY_CLASS.has(name)) return "A";
  if (name === "IF") return i === 0 ? "V" : "self";
  if (name === "CHOOSE") return i === 0 ? "V" : "self";
  if (name === "OFFSET") return i === 0 ? "R" : "V";
  if (name === "INDIRECT") return "V";
  if (name === "SUBTOTAL") return i === 0 ? "V" : "R";
  if (name === "AGGREGATE") return i < 2 ? "V" : "R";
  if (RANGE_ARGS.has(name)) return "R";
  return takesOneAt(name, nargs, i) ? "V" : "R";
}

const LEVELS: readonly (readonly string[])[] = [["=", "<>", "<", ">", "<=", ">="], ["&"], ["+", "-"], ["*", "/"], ["^"]];

class Reader {
  readonly #toks: Token[];
  #i = 0;
  readonly #names: ReadonlySet<string>;
  constructor(toks: Token[], names: ReadonlySet<string>) { this.#toks = toks; this.#names = names; }

  // The @s already written: where each starts and where what it asks for ends.
  readonly ats: [number, number][] = [];

  #cur(): Token { return this.#toks[this.#i] ?? { kind: "end", at: 0, text: "" }; }

  done(): boolean { return this.#cur().kind === "end"; }

  expr(level = 0): Node {
    if (level === LEVELS.length) return this.#unary();
    let left = this.expr(level + 1);
    for (let t = this.#cur(); t.kind === "op" && LEVELS[level].includes(t.op); t = this.#cur()) {
      this.#i++;
      left = { t: "bin", left, right: this.expr(level + 1) };
    }
    return left;
  }

  #unary(): Node {
    const t = this.#cur();
    if (t.kind === "op" && (t.op === "-" || t.op === "+")) { this.#i++; return { t: "un", arg: this.#unary() }; }
    // An @ already there is the one value it asks for.
    if (t.kind === "@") {
      this.#i++;
      this.#unary();
      const last = this.#toks[this.#i - 1];
      this.ats.push([t.at, last.at + last.text.length]);
      return { t: "other" };
    }
    let e = this.#primary();
    while (this.#cur().kind === "%") { this.#i++; e = { t: "un", arg: e }; }
    return e;
  }

  #primary(): Node {
    const t = this.#cur();
    this.#i++;
    switch (t.kind) {
      case "wide": return { t: "range", start: t.at };
      case "table": return /@|#This Row/i.test(t.spec) ? { t: "other" } : { t: "range", start: t.at };
      case "ref": {
        // A1:B2 (a range of one cell is that cell), A1# (what it spills).
        if (this.#cur().kind === ":" && this.#toks[this.#i + 1]?.kind === "ref") {
          const to = this.#toks[this.#i + 1] as Extract<Token, { kind: "ref" }>;
          this.#i += 2;
          return t.name === to.name ? { t: "other" } : { t: "range", start: t.at };
        }
        return t.spill ? { t: "range", start: t.at } : { t: "other" };
      }
      case "name": {
        if (this.#cur().kind !== "(") return this.#names.has(t.name) ? { t: "range", start: t.at } : { t: "other" };
        this.#i++;
        const args: Node[] = [];
        if (this.#cur().kind !== ")") {
          for (;;) {
            args.push(this.#cur().kind === "," || this.#cur().kind === ")" ? { t: "other" } : this.expr());
            if (this.#cur().kind !== ",") break;
            this.#i++;
          }
        }
        if (this.#cur().kind !== ")") throw new RangeError("a call never closed");
        this.#i++;
        // What these call is not an expression of their own (LET, LAMBDA, SINGLE: one value).
        if (t.name === "LET" || t.name === "LAMBDA" || t.name === "SINGLE") return { t: "other" };
        // A call of a call, f(1)(2): left alone.
        if (this.#cur().kind === "(") throw new RangeError("a call of a call");
        return { t: "call", start: t.at, name: t.name, args };
      }
      case "(": {
        const inner = this.expr();
        if (this.#cur().kind !== ")") throw new RangeError("a ( never closed");
        this.#i++;
        return { t: "paren", inner };
      }
      case "{":
        while (this.#cur().kind !== "}" && this.#cur().kind !== "end") this.#i++;
        this.#i++;
        return { t: "other" };
      default:
        return { t: "other" };
    }
  }
}

// `scalar`: where one value goes (an operand, an argument of one value). `array`: inside what an
// array-class argument works out, where nothing is intersected.
function walk(n: Node, array: boolean, one: boolean, at: Set<number>): void {
  if (array) return;
  switch (n.t) {
    case "range": if (one) at.add(n.start); return;
    case "bin": walk(n.left, false, true, at); walk(n.right, false, true, at); return;
    case "un": walk(n.arg, false, true, at); return;
    case "paren": walk(n.inner, false, one, at); return;
    case "call": {
      if (one && (MAY_GIVE_RANGE.has(n.name) || ((n.name === "ROW" || n.name === "COLUMN") && n.args.length === 1 && n.args[0].t === "range"))) at.add(n.start);
      n.args.forEach((a, i) => {
        const c = classOf(n.name, n.args.length, i);
        if (c === "A") return;
        walk(a, false, c === "V" || (c === "self" && one), at);
      });
      return;
    }
    default: return;
  }
}

// The names (upper case) of a workbook's defined names that are a range of more than one cell.
export type RangeNames = ReadonlySet<string>;

// The formula (what follows the '=') as Excel's loader shows it when the file gave it no array flag.
export function intersected(formula: string, names: RangeNames = new Set()): string {
  const at = insertions(formula, names);
  let out = formula;
  for (const i of [...at].sort((a, b) => b - a)) out = out.slice(0, i) + "@" + out.slice(i);
  return out;
}

function insertions(formula: string, names: RangeNames): Set<number> {
  const at = new Set<number>();
  try {
    const r = new Reader([...tokens(formula, 0)], names);
    const root = r.expr();
    if (r.done()) walk(root, false, true, at);
  } catch {
    // A formula that does not read is left as it is.
  }
  return at;
}

// The formula with every @ taken out (those written, and so those the loader would put in).
function plain(formula: string): string {
  let out = formula;
  try {
    for (const t of [...tokens(formula, 0)].reverse()) if (t.kind === "@") out = out.slice(0, t.at) + out.slice(t.at + 1);
  } catch {
    return formula;
  }
  return out;
}

// The names (upper case) among a workbook's and a sheet's defined names (name, what it is) that are
// a range of more than one cell.
export function rangeNamesOf(defined: readonly (readonly [string, string])[]): Set<string> {
  const out = new Set<string>();
  for (const [n, r] of defined) {
    const m = /^(?:.*!)?\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(r.trim());
    if (m && (m[1].toUpperCase() !== m[3].toUpperCase() || m[2] !== m[4])) out.add(n.toUpperCase());
  }
  return out;
}

// Each @ written as SINGLE(...), the name a file keeps for it.
export function singled(formula: string): string {
  // Each @ becomes SINGLE( and what it asks for ends with ); from the end, so places stay true.
  const edits: [number, number, string][] = [];
  try {
    const r = new Reader([...tokens(formula, 0)], new Set());
    r.expr();
    for (const [a, b] of r.ats) edits.push([a, 1, "SINGLE("], [b, 0, ")"]);
  } catch {
    return formula;
  }
  edits.sort((x, y) => (y[0] - x[0]) || (x[1] - y[1]));
  let out = formula;
  for (const [i, drop, s] of edits) out = out.slice(0, i) + s + out.slice(i + drop);
  return out;
}

// The formula as a file keeps it: `legacy` when the loader would put in exactly its @s (it is
// then written with none, and no array flag), else every @ as SINGLE(...) and the formula an
// array formula (a dynamic one's flag is the array's), which the loader leaves as it is.
export function stored(formula: string, names: RangeNames = new Set()): { readonly text: string; readonly legacy: boolean } {
  const bare = plain(formula);
  return intersected(bare, names) === formula ? { text: bare, legacy: true } : { text: singled(formula), legacy: false };
}
