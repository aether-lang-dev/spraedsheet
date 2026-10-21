// Formula compatibility workbooks from IronCalc, group "statistical" (docs/comparisons/IRONCALC.md
// action 2; third_party/ironcalc/PROVENANCE.md): workbooks Excel saved, with the value it
// worked out stored beside each formula. Each is opened here, and every formula cell must
// come out as Excel's stored value, but for the cells listed in lib/ironcalc_calc_known.ts
// (with how they differ): a listed cell that comes right is reported, so the list shrinks
// as the functions are fixed. The assertions are ours; the workbooks are IronCalc's, byte
// for byte. (tests/headless/lib/mkironcalc_calc.py writes this file and its fixtures.)
import { FIXTURES } from "./lib/ironcalc_calc_statistical.ts";
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

// expect: ok STATISTICAL/AVEDEV.xlsx
// expect: ok STATISTICAL/AVERAGE.xlsx
// expect: ok STATISTICAL/BETA_GAMMA.xlsx
// expect: ok STATISTICAL/CHISQ.xlsx
// expect: ok STATISTICAL/CHISQ_TEST.xlsx
// expect: ok STATISTICAL/CONFIDENCE.xlsx
// expect: ok STATISTICAL/CORREL_SLOPE_INTERCEPT_RSQ_STEYX.xlsx
// expect: ok STATISTICAL/COUNT.xlsx
// expect: ok STATISTICAL/COVARIANCE.xlsx
// expect: ok STATISTICAL/DEVSQ.xlsx
// expect: ok STATISTICAL/EXPONDIST.xlsx
// expect: ok STATISTICAL/EXPON_DIST.xlsx
// expect: ok STATISTICAL/FISHER.xlsx
// expect: ok STATISTICAL/FISHER_FISHERINV.xlsx
// expect: ok STATISTICAL/FREQUENCY.xlsx
// expect: ok STATISTICAL/F_DIST.xlsx
// expect: ok STATISTICAL/F_TEST.xlsx
// expect: ok STATISTICAL/GAMMALN_GAMMALNPRECISE.xlsx
// expect: ok STATISTICAL/GAUSS.xlsx
// expect: ok STATISTICAL/HYPERGEOM_DIST.xlsx
// expect: ok STATISTICAL/LINEST_LOGEST.xlsx
// expect: ok STATISTICAL/LOGNORM.xlsx
// expect: ok STATISTICAL/MEADIAN_KURT_SKEW_HARMEAN.xlsx
// expect: ok STATISTICAL/MINA_MAXA.xlsx
// expect: ok STATISTICAL/MIN_MAX.xlsx
// expect: ok STATISTICAL/MODE.xlsx
// expect: ok STATISTICAL/NORM_DIST.xlsx
// expect: ok STATISTICAL/PEARSON.xlsx
// expect: ok STATISTICAL/PERCENTIL_QUARTIL.xlsx
// expect: ok STATISTICAL/PERMUT_PERMUTATIONA.xlsx
// expect: ok STATISTICAL/PHI.xlsx
// expect: ok STATISTICAL/POISSON.xlsx
// expect: ok STATISTICAL/PROB.xlsx
// expect: ok STATISTICAL/RANK_EQ_RANK_AVG.xlsx
// expect: ok STATISTICAL/SMALL_LARGE.xlsx
// expect: ok STATISTICAL/STANDARIZE.xlsx
// expect: ok STATISTICAL/STDEV.xlsx
// expect: ok STATISTICAL/TREND_GROWTH.xlsx
// expect: ok STATISTICAL/TRIMMEAN.xlsx
// expect: ok STATISTICAL/T_DIST.xlsx
// expect: ok STATISTICAL/T_TEST.xlsx
// expect: ok STATISTICAL/VARIANCE.xlsx
// expect: ok STATISTICAL/WEIBULL_DIST.xlsx
// expect: ok STATISTICAL/Z_TEST.xlsx
// expect: ok every workbook was opened and its formulas compared
