// Part three of the original-workbook corpus round trip (lib/corpus_roundtrip.ts; docs/comparisons/IRONCALC.md action 5):
// IronCalc's Excel-authored workbooks of the groups financial, lookup_and_reference, statistical, text, opened, edited, saved as XLSX and reopened, the meaning the same.
// The corpus is split so that each part fits the 60 s a test has on a slow host (cachyos runs these about four times slower than the Mac).
import { FIXTURES as financial } from "./lib/ironcalc_calc_financial.ts";
import { FIXTURES as lookup_and_reference } from "./lib/ironcalc_calc_lookup_and_reference.ts";
import { FIXTURES as statistical } from "./lib/ironcalc_calc_statistical.ts";
import { FIXTURES as text } from "./lib/ironcalc_calc_text.ts";
import { roundtrip } from "./lib/corpus_roundtrip.ts";
import { eq } from "./lib/check.ts";

const { lost, files, checked } = await roundtrip([...financial, ...lookup_and_reference, ...statistical, ...text].filter(([n]) => true));
eq("every workbook of the corpus opens, saves and reopens to the same meaning", lost, []);
eq("and it was a corpus", [files >= 1, checked > 300], [true, true]);

// expect: ok every workbook of the corpus opens, saves and reopens to the same meaning
// expect: ok and it was a corpus
