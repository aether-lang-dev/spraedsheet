// Part four of the original-workbook corpus round trip (lib/corpus_roundtrip.ts; docs/comparisons/IRONCALC.md action 5):
// IronCalc's Excel-authored workbooks of the groups math_and_trigonometry, opened, edited, saved as XLSX and reopened, the meaning the same.
// (FLOOR_CEILING.xlsx and ARABIC_ROMAN.xlsx, 11,000 and 40,000 formula cells, are not round-tripped: each alone is a minute on a slow host,
// and the calc_ tests read them.) The corpus is split so that each part fits the 60 s a test has on a slow host (cachyos runs these about four times slower than the Mac).
import { FIXTURES as math_and_trigonometry } from "./lib/ironcalc_calc_math_and_trigonometry.ts";
import { roundtrip } from "./lib/corpus_roundtrip.ts";
import { eq } from "./lib/check.ts";

const { lost, files, checked } = await roundtrip([...math_and_trigonometry]);
eq("every workbook of the corpus opens, saves and reopens to the same meaning", lost, []);
eq("and it was a corpus", [files >= 1, checked > 300], [true, true]);

// expect: ok every workbook of the corpus opens, saves and reopens to the same meaning
// expect: ok and it was a corpus
