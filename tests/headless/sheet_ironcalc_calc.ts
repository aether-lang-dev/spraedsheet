// Formula compatibility workbooks from IronCalc (docs/comparisons/IRONCALC.md
// action 2; third_party/ironcalc/PROVENANCE.md): workbooks Excel saved, with the
// value it worked out stored beside each formula. Each is opened here, and every
// formula cell must come out as Excel's stored value, but for the cells listed
// in KNOWN (with why): a listed cell that comes right is reported, so the list
// shrinks as the functions are fixed. The assertions are ours; the workbooks are
// IronCalc's, byte for byte.
import { CALC_FIXTURES } from "./lib/ironcalc_calc_fixtures.ts";
import { disagreements } from "./lib/ironcalc_cached.ts";
import { eq } from "./lib/check.ts";

// workbook -> the cells known to differ from Excel, "Sheet!A1" (empty: none)
const KNOWN: Record<string, string[]> = {};

let total = 0;
for (const [name, b64] of CALC_FIXTURES) {
  const { diffs, checked } = await disagreements(name, b64);
  total += checked;
  const known = new Set(KNOWN[name] ?? []);
  const cell = (d: string) => d.slice(0, d.indexOf(": "));
  eq(name, [diffs.filter((d) => !known.has(cell(d))), [...known].filter((k) => !diffs.some((d) => cell(d) === k))], [[], []]);
}
eq("every workbook was opened and its formulas compared", total > 1000, true);

// expect: ok array_in_scalar_repro.xlsx
// expect: ok consume_string.xlsx
// expect: ok escape_strings.xlsx
// expect: ok issue_341.xlsx
// expect: ok simple_arrays.xlsx
// expect: ok DATABASE/DAVERAGE.xlsx
// expect: ok DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM_issue_1502.xlsx
// expect: ok ENGINEERING/ERFs.xlsx
// expect: ok ENGINEERING/GESTEP_DELTA.xlsx
// expect: ok FINANCIAL/COUPDAYS.xlsx
// expect: ok FINANCIAL/DISC.xlsx
// expect: ok FINANCIAL/FVSCHEDULE.xlsx
// expect: ok FINANCIAL/INTRATE.xlsx
// expect: ok FINANCIAL/NOMINAL_EFFECT.xlsx
// expect: ok FINANCIAL/NPV.xlsx
// expect: ok FINANCIAL/RRI.xlsx
// expect: ok INFORMATION/SHEET_SHEETS.xlsx
// expect: ok LOGICAL/BYCOL_BYROW.xlsx
// expect: ok LOGICAL/IFERROR.xlsx
// expect: ok LOGICAL/LET.xlsx
// expect: ok LOGICAL/MAKEARRAY.xlsx
// expect: ok LOGICAL/MAP_REDUCE.xlsx
// expect: ok LOGICAL/SCAN.xlsx
// expect: ok LOOKUP_AND_REFERENCE/ADDRESS.xlsx
// expect: ok LOOKUP_AND_REFERENCE/CHOOSECOLS_CHOOSEROWS.xlsx
// expect: ok LOOKUP_AND_REFERENCE/EXPAND.xlsx
// expect: ok LOOKUP_AND_REFERENCE/HSTACK_VSTACK.xlsx
// expect: ok LOOKUP_AND_REFERENCE/MATCH_LOOKUP_arrays.xlsx
// expect: ok LOOKUP_AND_REFERENCE/TOCOL.xlsx
// expect: ok LOOKUP_AND_REFERENCE/TRANSPOSE.xlsx
// expect: ok LOOKUP_AND_REFERENCE/WRAPCOLS_ROWS.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/BASE.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SUMIF_array.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SUMPRODUCT.xlsx
// expect: ok STATISTICAL/AVEDEV.xlsx
// expect: ok STATISTICAL/CHISQ_TEST.xlsx
// expect: ok STATISTICAL/CORREL_SLOPE_INTERCEPT_RSQ_STEYX.xlsx
// expect: ok STATISTICAL/DEVSQ.xlsx
// expect: ok STATISTICAL/FISHER.xlsx
// expect: ok STATISTICAL/F_TEST.xlsx
// expect: ok STATISTICAL/MIN_MAX.xlsx
// expect: ok STATISTICAL/MODE.xlsx
// expect: ok STATISTICAL/PEARSON.xlsx
// expect: ok STATISTICAL/PERCENTIL_QUARTIL.xlsx
// expect: ok STATISTICAL/PHI.xlsx
// expect: ok STATISTICAL/RANK_EQ_RANK_AVG.xlsx
// expect: ok STATISTICAL/STANDARIZE.xlsx
// expect: ok STATISTICAL/TRIMMEAN.xlsx
// expect: ok STATISTICAL/T_TEST.xlsx
// expect: ok STATISTICAL/Z_TEST.xlsx
// expect: ok TEXT/CONCAT.xlsx
// expect: ok TEXT/DOLLARs.xlsx
// expect: ok TEXT/EXACT.xlsx
// expect: ok TEXT/MID_LEFT_RIGHT_arrays.xlsx
// expect: ok TEXT/PROPER.xlsx
// expect: ok TEXT/REGEX.xlsx
// expect: ok TEXT/TEXTSPLIT.xlsx
// expect: ok TEXT/UNICODE.xlsx
// expect: ok TEXT/UPPER_LOWER.xlsx
// expect: ok every workbook was opened and its formulas compared
