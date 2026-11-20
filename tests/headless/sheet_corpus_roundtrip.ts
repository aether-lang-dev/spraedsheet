// Part one of the original-workbook corpus round trip (lib/corpus_roundtrip.ts; docs/comparisons/IRONCALC.md action 5):
// IronCalc's Excel-authored workbooks of the groups database, engineering, information, logical, opened, edited, saved as XLSX and reopened, the meaning the same.
// The corpus is split so that each part fits the 60 s a test has on a slow host (cachyos runs these about four times slower than the Mac).
import { FIXTURES as database } from "./lib/ironcalc_calc_database.ts";
import { FIXTURES as engineering } from "./lib/ironcalc_calc_engineering.ts";
import { FIXTURES as information } from "./lib/ironcalc_calc_information.ts";
import { FIXTURES as logical } from "./lib/ironcalc_calc_logical.ts";
import { roundtrip } from "./lib/corpus_roundtrip.ts";
import { eq } from "./lib/check.ts";

const { lost, files, checked } = await roundtrip([...database, ...engineering, ...information, ...logical].filter(([n]) => true));
eq("every workbook of the corpus opens, saves and reopens to the same meaning", lost, []);
eq("and it was a corpus", [files >= 1, checked > 300], [true, true]);

// expect: ok every workbook of the corpus opens, saves and reopens to the same meaning
// expect: ok and it was a corpus
