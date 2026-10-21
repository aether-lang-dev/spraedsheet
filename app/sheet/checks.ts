// Assertions in the sheet (blog-entry-todo.md B.6): a check is a formula
// that must be TRUE, with a name and a message, kept in the workbook and
// saved with its file. Worked out on every recalculation (Commands.execute
// after each command, a file when it is opened), a failing one is named in the
// status bar and listed in the Audit panel, with its message.
//
// A check passes when its formula gives TRUE, or a number that is not 0. FALSE,
// 0, text, an empty cell and an error all fail, and `shown` says which.
export interface Check {
  readonly name: string;
  // Without its leading "=".
  readonly formula: string;
  readonly message: string;
  // The sheet it is worked out on ("" is the first): the cells it names are that sheet's.
  readonly sheet: string;
}

export interface CheckResult {
  readonly name: string;
  readonly message: string;
  readonly ok: boolean;
  // What the formula gave, as a cell would show it.
  readonly shown: string;
}

// The status bar's words for the results: "" when all pass.
export function checksLine(results: readonly CheckResult[]): string {
  const failing = results.filter((r) => !r.ok);
  return failing.length ? `${failing.length} check${failing.length === 1 ? "" : "s"} failing: ${failing.map((r) => r.name).join(", ")}` : "";
}

// The checks of a file's JSON; one that is not a name, a formula and a message is left out.
export function checksOf(raw: unknown): Check[] {
  if (!Array.isArray(raw)) return [];
  return raw.filter((c: any) => c && typeof c.name === "string" && typeof c.formula === "string" && typeof c.message === "string" && c.name.trim() !== "" && c.formula.trim() !== "")
    .map((c: any): Check => ({ name: c.name, formula: c.formula, message: c.message, sheet: typeof c.sheet === "string" ? c.sheet : "" }));
}
