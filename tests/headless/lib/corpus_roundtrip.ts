// The corpus round trip of sheet_corpus_roundtrip_a and _b (docs/comparisons/IRONCALC.md action 5): each workbook is opened,
// worked out, edited, saved as XLSX, and reopened; what it means (each sheet's cell inputs, the values they come to) must
// be the same before and after, compared as meaning and not as ZIP bytes. The corpus is two tests, each with its own
// time limit, so a slow host (cachyos takes half as long again as the Mac) has a sensible budget for each.
import { Commands } from "../sheet/commands.ts";
import { Workbook } from "../sheet/workbook.ts";
import { writeXlsx } from "../sheet/xlsxwrite.ts";
import { toBase64 } from "../sheet/csv.ts";
import { show } from "../sheet/format.ts";

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

export async function roundtrip(corpus: readonly (readonly [string, string])[]): Promise<{ lost: string[]; files: number; checked: number }> {
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
return { lost, files, checked };
}
