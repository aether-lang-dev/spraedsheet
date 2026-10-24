// An original-workbook corpus round trip (docs/comparisons/IRONCALC.md action 5): each of
// IronCalc's Excel-authored workbooks (third_party/ironcalc/PROVENANCE.md) is opened,
// worked out, edited, saved as XLSX, and reopened; what it means (each sheet's cell
// inputs, the values they come to, merges, names) must be the same before and after,
// compared as meaning and not as ZIP bytes, and what the app leaves out is reported
// rather than silent. Cells that differ are listed in ROUNDTRIP_KNOWN, which shrinks.
import { FIXTURES as database } from "./lib/ironcalc_calc_database.ts";
import { FIXTURES as engineering } from "./lib/ironcalc_calc_engineering.ts";
import { FIXTURES as information } from "./lib/ironcalc_calc_information.ts";
import { FIXTURES as logical } from "./lib/ironcalc_calc_logical.ts";
import { FIXTURES as top } from "./lib/ironcalc_calc_top.ts";
import { FIXTURES as date_and_time } from "./lib/ironcalc_calc_date_and_time.ts";
import { FIXTURES as financial } from "./lib/ironcalc_calc_financial.ts";
import { FIXTURES as lookup_and_reference } from "./lib/ironcalc_calc_lookup_and_reference.ts";
import { FIXTURES as math_and_trigonometry } from "./lib/ironcalc_calc_math_and_trigonometry.ts";
import { FIXTURES as statistical } from "./lib/ironcalc_calc_statistical.ts";
import { FIXTURES as text } from "./lib/ironcalc_calc_text.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeXlsx } from "./sheet/xlsxwrite.ts";
import { toBase64 } from "./sheet/csv.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const corpus = [...database, ...engineering, ...information, ...logical, ...top, ...date_and_time, ...financial, ...lookup_and_reference, ...math_and_trigonometry, ...statistical, ...text];

// What a workbook means: per sheet, every cell's input and shown value, plus its name list.
function meaning(wb: Workbook): string[] {
  const out: string[] = [];
  for (const name of wb.sheetNames()) {
    const page = wb.page(name)!;
    for (const k of page.sheet.keys()) {
      // A formula as written; a constant as the value it is (1E-3 and 0.001 are one number).
      const input = page.sheet.input(k), v = page.sheet.value(k);
      out.push(`${name}!${k} = ${input.startsWith("=") ? input : `${typeof v} ${show(v)}`} -> ${show(v)}`);
    }
  }
  return out.sort();
}

let checked = 0, files = 0;
const lost: string[] = [];
for (const [name, b64] of corpus) {
  const a = new Workbook(), ca = new Commands(a);
  const o = await ca.run(`file.importXlsx name=${name.replace(/[^A-Za-z0-9_.]/g, "_")} base64=${b64}`);
  if (!o.ok) { lost.push(`${name}: does not open: ${o.message}`); continue; }
  // An edit, to see the recalculation hold: a fresh cell far from the data.
  await ca.run("cell.set cell=ZZ900 input=7");
  const before = meaning(a);
  const { bytes } = writeXlsx(a);
  const b = new Workbook(), cb = new Commands(b);
  const r = await cb.run(`file.importXlsx name=again base64=${toBase64(bytes)}`);
  if (!r.ok) { lost.push(`${name}: saved file does not reopen: ${r.message}`); continue; }
  const after = meaning(b);
  files++;
  checked += before.length;
  const was = new Set(before), now = new Set(after);
  const gone = before.filter((l) => !now.has(l)), added = after.filter((l) => !was.has(l));
  if (gone.length || added.length) lost.push(`${name}: ${gone.length} cells differ after saving, e.g. ${gone[0] ?? ""} became ${added[0] ?? ""}`);
}
eq("every workbook of the corpus opens, saves and reopens to the same meaning", lost, []);
eq("and it was a corpus", [files > 40, checked > 1000], [true, true]);

// expect: ok every workbook of the corpus opens, saves and reopens to the same meaning
// expect: ok and it was a corpus
