// Fill Series cases from IronCalc (MIT; third_party/ironcalc/PROVENANCE.md),
// run against app/sheet/series.ts as data. They are the inputs and expected
// results of the tests in base/src/user_model/sequence_detector.rs (English
// locale: its de/fr/es date and number cases are the locale work, not ours).
// IronCalc reads every selected cell and detects a progression across them,
// and gives none when they disagree; ours reads the first one or two cells
// (edit.fillSeries) and goes on from them. So each row is
//   [upstream test, the cells, what comes after them]
// with the cells cut to the first two (a third that IronCalc checks comes
// back as the first thing continued), and the rows where IronCalc finds no
// progression (a third cell that breaks the step) are not here: there the
// two designs differ on purpose.
import { continued } from "./sheet/series.ts";
import { eq } from "./lib/check.ts";

type Row = [string, string[], string[]];
const ROWS: Row[] = [
  ["test_numeric_progression_detector", ["1", "2"], ["3", "4", "5"]],
  ["test_numeric_progression_detector", ["10", "8"], ["6", "4", "2"]],
  ["test_numeric_progression_detector", ["-10", "-8"], ["-6", "-4", "-2"]],
  ["test_numeric_float", ["1.5", "2.0"], ["2.5", "3", "3.5"]],
  ["test_numeric_float", ["0.1", "0.2"], ["0.3", "0.4"]],
  ["test_numeric_float", ["0.001", "0.002"], ["0.003", "0.004"]],
  ["test_numeric_float", ["-1.5", "-1.0"], ["-0.5", "0", "0.5"]],
  ["test_numeric_float", ["10.5", "9.5"], ["8.5", "7.5"]],
  ["test_numeric_float", ["10000.4000000007", "10000.4000000008"], ["10000.4000000009", "10000.400000001"]],
  ["test_numeric_float", ["0.3", "0.6"], ["0.9", "1.2"]],
  ["test_numeric_grouping_validation", ["1000000", "2000000"], ["3000000"]],
  ["test_numeric_grouping_validation", ["1000.50", "2000.50"], ["3000.5"]],
  ["test_suffixed_progression_detector", ["A1", "A2"], ["A3", "A4"]],
  ["test_suffixed_progression_detector", ["A0.1", "A0.2"], ["A0.3", "A0.4", "A0.5", "A0.6", "A0.7", "A0.8", "A0.9", "A0.10"]],
  ["test_suffixed_progression_detector", ["Product 1", "Product 2"], ["Product 3"]],
  ["test_suffixed_progression_detector", ["Q10", "Q9"], ["Q8"]],
  ["test_suffixed_progression_float_like_suffix", ["V1.0", "V1.5"], ["V1.10"]],
  ["test_detect_progression", ["1", "3"], ["5"]],
  ["test_detect_progression", ["X10", "X20"], ["X30"]],
  ["test_detect_progression", ["Mar", "Apr"], ["May"]],
  ["test_date_progression_detector_en", ["Monday", "Tuesday"], ["Wednesday"]],
  ["test_date_progression_detector_en", ["Mon", "Tue"], ["Wed"]],
  ["test_date_progression_detector_en", ["January", "February"], ["March"]],
  ["test_date_progression_detector_en", ["Jan", "Feb"], ["Mar"]],
  ["test_date_progression_detector_en", ["Jan ", "Feb "], ["Mar"]],
  ["test_date_progression_detector_en", ["Saturday", "Sunday"], ["Monday"]],
  ["test_date_progression_detector_en", ["Jan", "Mar"], ["May"]],
  ["test_date_progression_detector_en", ["jan", "feb"], ["mar"]],
  ["test_date_progression_detector_en", ["saturday", "SUNDAY"], ["monday"]],
  ["test_date_progression_detector_en", ["monday", "tuesday"], ["wednesday"]],
  ["test_date_progression_detector_en", ["january", "february"], ["march"]],
  ["test_date_progression_detector_en", ["MONDAY", "TUESDAY"], ["WEDNESDAY"]],
  ["test_date_progression_detector_en", ["JANUARY", "FEBRUARY"], ["MARCH"]],
  ["test_date_progression_detector_en", ["JAN", "FEB"], ["MAR"]],
  // no progression: copied
  ["test_suffixed_progression_detector", ["A1", "B2"], ["A1", "B2", "A1"]],
  ["test_suffixed_progression_detector", ["Test-A", "Test-B"], ["Test-A", "Test-B", "Test-A"]],
  ["test_suffixed_progression_detector", ["Test", "Test"], ["Test", "Test"]],
];
const diffs: string[] = [];
for (const [name, cells, want] of ROWS) {
  const got = continued(cells, want.length);
  if (JSON.stringify(got) !== JSON.stringify(want)) diffs.push(`${name} ${cells.join("|")}: got ${got.join("|")}, want ${want.join("|")}`);
}
eq("ironcalc series: the upstream cases", diffs, []);
// expect: ok ironcalc series: the upstream cases
