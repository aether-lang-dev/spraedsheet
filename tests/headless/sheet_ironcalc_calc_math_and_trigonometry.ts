// Formula compatibility workbooks from IronCalc, group "math_and_trigonometry" (docs/comparisons/IRONCALC.md
// action 2; third_party/ironcalc/PROVENANCE.md): workbooks Excel saved, with the value it
// worked out stored beside each formula. Each is opened here, and every formula cell must
// come out as Excel's stored value, but for the cells listed in lib/ironcalc_calc_known.ts
// (with how they differ): a listed cell that comes right is reported, so the list shrinks
// as the functions are fixed. The assertions are ours; the workbooks are IronCalc's, byte
// for byte. (tests/headless/lib/mkironcalc_calc.py writes this file and its fixtures.)
import { FIXTURES } from "./lib/ironcalc_calc_math_and_trigonometry.ts";
import { KNOWN } from "./lib/ironcalc_calc_known.ts";
import { disagreements } from "./lib/ironcalc_cached.ts";
import { eq } from "./lib/check.ts";

let total = 0;
for (const [name, b64] of FIXTURES) {
  const { diffs, checked } = await disagreements(name, b64);
  total += checked;
  const known = KNOWN[name] ?? {};
  const now = new Map(diffs.map((d): [string, string] => [d.cell, d.kind]));
  // fresh: differs and not listed; gone: listed and now right; changed: listed, differs another way
  const fresh = [...now].filter(([c]) => !(c in known)).map(([c, k]) => `${c} (${k})`);
  const gone = Object.keys(known).filter((c) => !now.has(c));
  const changed = [...now].filter(([c, k]) => c in known && known[c] !== k).map(([c, k]) => `${c} (${known[c]} -> ${k})`);
  eq(name, [fresh, gone, changed], [[], [], []]);
}
eq("every workbook was opened and its formulas compared", total > 0, true);

// expect: ok MATH_AND_TRIGONOMETRY/ABS.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/ATAN2_POWER.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/BASE.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/COMBIN_COMBINA.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/DEGREES_RADIANS.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/EVEN_ODD.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/EXP_SIGN.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/FACT_DOUBLEFACT.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/LOG_LOG10_LN.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/MMULT.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/MOD_QUOTIENT.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/MULTINOMIAL.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/MUNIT_MDETERM_MINVERSE.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/POWER.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/PRODUCT.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/PRODUCT_SUM.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SEQUENCE.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SERIESSUM.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SQRT_SQRTPI.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SUMIF_AVERAGE_IF.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SUMIF_array.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SUMPRODUCT.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SUMSQ.xlsx
// expect: ok MATH_AND_TRIGONOMETRY/SUMX2MY2_SUMX2PY2_SUMXMY2.xlsx
// expect: ok every workbook was opened and its formulas compared
