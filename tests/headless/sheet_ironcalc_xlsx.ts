// Two XLSX regressions from IronCalc (MIT; third_party/ironcalc/PROVENANCE.md),
// their workbooks copied unchanged to third_party/ironcalc/fixtures/xlsx/ and
// read here as base64 (lib/ironcalc_fixtures.ts, made by lib/mkironcalc.py).
// The assertions are IronCalc's, translated: ours are the cases beside them.
import { readXlsx } from "./sheet/xlsx.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { SHARED_FORMULA_VOLATILE, OPTIONAL_XF_ID } from "./lib/ironcalc_fixtures.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

// From xlsx/tests/test_import_shared_formula_volatile.rs,
// test_import_volatile_shared_formula_daughter_cell: a daughter cell of a
// volatile shared formula is an empty <f t="shared" ca="1" si="1"/>; it must
// not be taken for a volatile spill placeholder (<f ca="1"/>), which would
// drop the formula. C9 is the master (ref C9:D9), D9 the daughter.
{
  const wb = new Workbook(), cmds = new Commands(wb);
  const o = await cmds.run(`file.importXlsx name=Vol.xlsx base64=${SHARED_FORMULA_VOLATILE}`);
  eq("ironcalc volatile shared formula: opened with nothing left out", [o.ok, readXlsx(fromBase64(SHARED_FORMULA_VOLATILE)).report], [true, []]);
  eq("ironcalc volatile shared formula: the master and the daughter keep their formulas", [wb.sheet.input("C9"), wb.sheet.input("D9")], ["=SUM(C3:C8)", "=SUM(D3:D8)"]);
  eq("ironcalc volatile shared formula: the daughter's value", [show(wb.sheet.value("C9")), show(wb.sheet.value("D9"))], ["21", "210"]);
}

// From xlsx/tests/optional_xf_id_import.rs, test_cell_xfs_without_xf_id_loads:
// xfId is optional on a cellXfs <xf> (many exports omit it); the file must
// load, and A2's style (vertical="top" wrapText="1") must still be read.
// Ours differs on the vertical alignment: the app has no vertical-alignment
// look, so only the wrap is asserted (see PROVENANCE.md, deliberate differences).
{
  const wb = new Workbook(), cmds = new Commands(wb);
  const o = await cmds.run(`file.importXlsx name=Xf.xlsx base64=${OPTIONAL_XF_ID}`);
  eq("ironcalc optional xfId: loads", o.ok, true);
  eq("ironcalc optional xfId: A2's wrap survives", wb.notes.get("A2")?.look?.wrap, true);
  eq("ironcalc optional xfId: the title row's merge, the frozen pane", [wb.page(wb.pages()[0].name)!.merges.length, wb.page(wb.pages()[0].name)!.frozen], [1, 2]);
}

// expect: ok ironcalc volatile shared formula: opened with nothing left out
// expect: ok ironcalc volatile shared formula: the master and the daughter keep their formulas
// expect: ok ironcalc volatile shared formula: the daughter's value
// expect: ok ironcalc optional xfId: loads
// expect: ok ironcalc optional xfId: A2's wrap survives
// expect: ok ironcalc optional xfId: the title row's merge, the frozen pane
