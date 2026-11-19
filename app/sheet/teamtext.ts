// A workbook's text form (textform.ts) as keyed lines, the unit a team
// workbook's operations change (docs/design/teamware-over-http.md, section 2).
// Pure, and importing nothing, so the server's tests (node) read the same
// module the page does.
//
// A key names what a line (or a few lines) is about, its parts joined by tabs
// (the text form's fields never hold a raw tab):
//
//   "Sheet1\tB2"            a cell of the sheet whose @sheet line names it Sheet1
//   "Sheet1\t@names\tQTY"   an entry of a sheet's object part
//   "Sheet1\t@view"         a sheet's setting (one value on its line)
//   "Sheet1\t@rules[]"      a sheet's list part, all its elements, in order
//   "\t@formulas"           the workbook's own (the sheet part is empty)
//   "\t@sheet"              the order of the sheets: their @sheet lines
//
// The value is the line itself (a list's, its lines joined with "\n"), so a
// key's value written out under its sheet is the text again. The sheet part
// is the name as its @sheet line spells it (quoted there when it must be).
// sheet-server (server/sheetserver.ae) keeps the same keys and writes the
// same text (textOfLines); tests/server/ holds the two to each other.

export const TEXT_HEADER = "# rexxsheet text 1";
export const SHEETS = "\t@sheet";

export interface Change {
  readonly key: string;
  readonly was: string;   // "" where there was no such line
  readonly now: string;   // "" where there is none now
}

// The text's lines by key.
export function keyedLines(text: string): Map<string, string> {
  const out = new Map<string, string>();
  const add = (key: string, line: string, many: boolean) => {
    const had = out.get(key);
    out.set(key, many && had !== undefined ? `${had}\n${line}` : line);
  };
  let sheet = "";
  for (const raw of text.split("\n")) {
    const line = raw.endsWith("\r") ? raw.slice(0, -1) : raw;
    if (line === "" || line.startsWith("#")) continue;
    const f = line.split("\t");
    if (f[0] === "@sheet") {
      sheet = f[1] ?? "";
      add(SHEETS, line, true);
    } else if (f[0].startsWith("@")) {
      if (f[0].endsWith("[]")) add(`${sheet}\t${f[0]}`, line, true);
      else if (f.length >= 3) add(`${sheet}\t${f[0]}\t${f[1]}`, line, false);
      else add(`${sheet}\t${f[0]}`, line, false);
    } else add(`${sheet}\t${f[0]}`, line, false);
  }
  return out;
}

// The sheet a key belongs to ("" for the workbook's own, and for the order
// of the sheets).
export const sheetOfKey = (key: string): string => key.slice(0, key.indexOf("\t"));

// What changed from one set of lines to another, by key, in key order.
export function changesBetween(before: ReadonlyMap<string, string>, after: ReadonlyMap<string, string>): Change[] {
  const out: Change[] = [];
  for (const [key, now] of after) {
    const was = before.get(key) ?? "";
    if (was !== now) out.push({ key, was, now });
  }
  for (const [key, was] of before) if (!after.has(key)) out.push({ key, was, now: "" });
  return out.sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
}

// The first change whose `was` is not what `lines` holds now (the
// compare-and-set the server makes), or null when all of them hold.
export function staleChange(lines: ReadonlyMap<string, string>, changes: readonly Change[]): Change | null {
  for (const c of changes) if ((lines.get(c.key) ?? "") !== c.was) return c;
  return null;
}

// `changes` made to `lines` (in place), forwards or taken back.
export function applyChanges(lines: Map<string, string>, changes: readonly Change[], backwards = false): void {
  for (const c of changes) {
    const v = backwards ? c.was : c.now;
    if (v === "") lines.delete(c.key);
    else lines.set(c.key, v);
  }
}

// The changes taken back: each one's was and now swapped.
export const inverse = (changes: readonly Change[]): Change[] => changes.map((c) => ({ key: c.key, was: c.now, now: c.was }));

// A cell's column and row ("B12": 2, 12), or null for a part.
function cellOf(name: string): [number, number] | null {
  const m = /^([A-Z]{1,3})([0-9]{1,7})$/.exec(name);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]), col];
}

const byText = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);

// The text the lines are: the header, `# revision <rev>` when one is given,
// the workbook's own lines (by key), then each sheet in order, its cells by
// row and column, then its parts (by key). The server writes the same.
export function textOfLines(lines: ReadonlyMap<string, string>, rev?: number): string {
  const out: string[] = [TEXT_HEADER];
  if (rev !== undefined) out.push(`# revision ${rev}`);
  const bySheet = new Map<string, string[]>();
  for (const key of lines.keys()) {
    if (key === SHEETS) continue;
    const s = sheetOfKey(key);
    const list = bySheet.get(s);
    if (list) list.push(key);
    else bySheet.set(s, [key]);
  }
  for (const key of (bySheet.get("") ?? []).sort(byText)) out.push(lines.get(key)!);
  for (const sheetLine of (lines.get(SHEETS) ?? "").split("\n")) {
    if (sheetLine === "") continue;
    out.push(sheetLine);
    const name = sheetLine.slice(sheetLine.indexOf("\t") + 1);
    const keys = bySheet.get(name) ?? [];
    const cells: [number, number, string][] = [], parts: string[] = [];
    for (const key of keys) {
      const rest = key.slice(name.length + 1), at = cellOf(rest);
      if (at) cells.push([at[0], at[1], key]);
      else parts.push(key);
    }
    cells.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    for (const c of cells) out.push(lines.get(c[2])!);
    for (const key of parts.sort(byText)) out.push(lines.get(key)!);
  }
  return `${out.join("\n")}\n`;
}
