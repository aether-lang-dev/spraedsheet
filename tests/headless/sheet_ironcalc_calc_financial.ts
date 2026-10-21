// Formula compatibility workbooks from IronCalc, group "financial" (docs/comparisons/IRONCALC.md
// action 2; third_party/ironcalc/PROVENANCE.md): workbooks Excel saved, with the value it
// worked out stored beside each formula. Each is opened here, and every formula cell must
// come out as Excel's stored value, but for the cells listed in lib/ironcalc_calc_known.ts
// (with how they differ): a listed cell that comes right is reported, so the list shrinks
// as the functions are fixed. The assertions are ours; the workbooks are IronCalc's, byte
// for byte. (tests/headless/lib/mkironcalc_calc.py writes this file and its fixtures.)
import { FIXTURES } from "./lib/ironcalc_calc_financial.ts";
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

// expect: ok FINANCIAL/ACCRINT.xlsx
// expect: ok FINANCIAL/ACCRINTM.xlsx
// expect: ok FINANCIAL/COUPDAYS.xlsx
// expect: ok FINANCIAL/DB_DDB.xlsx
// expect: ok FINANCIAL/DISC.xlsx
// expect: ok FINANCIAL/FV.xlsx
// expect: ok FINANCIAL/FVSCHEDULE.xlsx
// expect: ok FINANCIAL/INTRATE.xlsx
// expect: ok FINANCIAL/IPMT_PPMT.xlsx
// expect: ok FINANCIAL/ISPMT.xlsx
// expect: ok FINANCIAL/NOMINAL_EFFECT.xlsx
// expect: ok FINANCIAL/NPER.xlsx
// expect: ok FINANCIAL/NPV.xlsx
// expect: ok FINANCIAL/PDURATION.xlsx
// expect: ok FINANCIAL/PMT.xlsx
// expect: ok FINANCIAL/PV.xlsx
// expect: ok FINANCIAL/RATE.xlsx
// expect: ok FINANCIAL/RRI.xlsx
// expect: ok FINANCIAL/SLN_SYD.xlsx
// expect: ok FINANCIAL/XNPV.xlsx
// expect: ok every workbook was opened and its formulas compared
