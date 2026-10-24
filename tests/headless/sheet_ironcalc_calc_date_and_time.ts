// Formula compatibility workbooks from IronCalc, group "date_and_time" (docs/comparisons/IRONCALC.md
// action 2; third_party/ironcalc/PROVENANCE.md): workbooks Excel saved, with the value it
// worked out stored beside each formula. Each is opened here, and every formula cell must
// come out as Excel's stored value, but for the cells listed in lib/ironcalc_calc_known.ts
// (with how they differ): a listed cell that comes right is reported, so the list shrinks
// as the functions are fixed. The assertions are ours; the workbooks are IronCalc's, byte
// for byte. (tests/headless/lib/mkironcalc_calc.py writes this file and its fixtures.)
import { FIXTURES } from "./lib/ironcalc_calc_date_and_time.ts";
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

// expect: ok DATE_AND_TIME/DATEVALUE_issue_1379.xlsx
// expect: ok DATE_AND_TIME/DATEVALUE_issue_1380.xlsx
// expect: ok DATE_AND_TIME/DATE_DAY_MONTH_YEAR.xlsx
// expect: ok DATE_AND_TIME/EDATE.xlsx
// expect: ok DATE_AND_TIME/EOMONTH.xlsx
// expect: ok DATE_AND_TIME/EOMONTH_issue_1388.xlsx
// expect: ok DATE_AND_TIME/TIMEVALUE.xlsx
// expect: ok DATE_AND_TIME/TIMEVALUE_issue_1402.xlsx
// expect: ok DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM_issue_1502.xlsx
// expect: ok DATE_AND_TIME/YEARFRAC_DATEDIF.xlsx
// expect: ok every workbook was opened and its formulas compared
