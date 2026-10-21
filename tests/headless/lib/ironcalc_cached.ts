// The values Excel stored beside the formulas of a workbook (<v> of each <c>
// with an <f>), to compare what this app works out for them: the
// IronCalc calc_tests workbooks were saved by Excel, so those cached values are
// an oracle for functions (third_party/ironcalc/PROVENANCE.md).
import { unzip, utf8Text } from "../sheet/zip.ts";
import { parseXml, all, kid, kids, attr, textOf } from "../sheet/xml.ts";
import { fromBase64 } from "../sheet/csv.ts";
import { Commands } from "../sheet/commands.ts";
import { Workbook } from "../sheet/workbook.ts";
import { show } from "../sheet/format.ts";
import { CellError } from "../sheet/values.ts";

// sheet name -> [cell, "type:cached text"] for each formula cell with a cached value
export function cachedValues(bytes: Uint8Array): Map<string, [string, string][]> {
  const parts = unzip(bytes);
  const wb = parseXml(utf8Text(parts.get("xl/workbook.xml")!));
  const relsXml = parseXml(utf8Text(parts.get("xl/_rels/workbook.xml.rels")!));
  const rel = new Map<string, string>();
  for (const r of all(relsXml, "Relationship")) rel.set(attr(r, "Id") ?? "", "xl/" + (attr(r, "Target") ?? "").replace(/^\/?(xl\/)?/, ""));
  const out = new Map<string, [string, string][]>();
  for (const s of kids(kid(wb, "sheets")!, "sheet")) {
    const xml = parts.get(rel.get(attr(s, "id") ?? "") ?? "");
    if (!xml) continue;
    const cells: [string, string][] = [];
    for (const c of all(parseXml(utf8Text(xml)), "c")) {
      const f = kid(c, "f"), v = kid(c, "v");
      if (f && v) cells.push([attr(c, "r") ?? "", `${attr(c, "t") ?? "n"}:${textOf(v)}`]);
    }
    out.set(attr(s, "name") ?? "", cells);
  }
  return out;
}

// The cells of a workbook (opened here from its base64) whose value is not what
// Excel cached, as "Sheet!A1: want n:3 got 4", and how many cells were compared.
export async function disagreements(name: string, b64: string): Promise<{ diffs: string[]; checked: number }> {
  const wb = new Workbook(), cmds = new Commands(wb);
  const o = await cmds.run(`file.importXlsx name=${name.replace(/[^A-Za-z0-9_.]/g, "_")} base64=${b64}`);
  if (!o.ok) return { diffs: [`import: ${o.message}`], checked: 0 };
  const diffs: string[] = [];
  let checked = 0;
  for (const [sheet, cells] of cachedValues(fromBase64(b64))) {
    const page = wb.page(sheet);
    if (!page) { diffs.push(`${sheet}: no such sheet`); continue; }
    for (const [ref, tv] of cells) {
      checked++;
      const i = tv.indexOf(":"), t = tv.slice(0, i), v = tv.slice(i + 1);
      const key = ref.replace(/\$/g, "");
      const got = page.sheet.value(key);
      let ok: boolean, shown: string;
      if (t === "e") { ok = got instanceof CellError && got.code === v; shown = got instanceof CellError ? got.code : show(got); }
      else if (t === "b") { ok = got === (v === "1"); shown = show(got); }
      else if (t === "str" || t === "s") { ok = typeof got === "string" && got === v; shown = show(got); }
      else { const n = Number(v); ok = typeof got === "number" && Math.abs(got - n) <= 1e-9 * Math.max(1, Math.abs(n)); shown = show(got); }
      if (!ok) diffs.push(`${sheet}!${key}: want ${t}:${v} got ${shown}`);
    }
  }
  return { diffs, checked };
}
