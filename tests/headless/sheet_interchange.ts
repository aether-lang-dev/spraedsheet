// Interchange honesty (docs/comparisons/LIBREOFFICE.md action 5): file.interchange publishes, for
// each format, whether it is read and written, what it keeps and what it loses,
// and the policy for a RexxJS workbook, which other applications cannot run:
// its results are written, not its formulas. The losses stated are the
// reports the writers actually give, checked here.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
type Row = { format: string; read: boolean; write: boolean; keeps: string; loses: string[] };
const o = await cmds.execute("file.interchange");
const rows = o.output as Row[];
eq("each format: read, written", rows.map((r) => [r.format, r.read, r.write]),
  [["xlsx", true, true], ["xls", true, false], ["ods", true, false], ["csv", true, true], ["json", true, true], ["html", true, true], ["pdf", false, true]]);
eq("ODS is read as cached values; XLS formulas it cannot write stay values", [rows.find((r) => r.format === "ods")!.loses.some((l) => l.includes("cached value")), rows.find((r) => r.format === "xls")!.loses.some((l) => l.includes("cached"))], [true, true]);
eq("the message", o.message, "7 formats; a RexxJS workbook is written to XLSX as its values, not its formulas");

// The XLSX losses are what the writer reports.
await wb.sheet.setMany([["A1", "1"], ["A2", "2"], ["B1", "=SPARKLINE(A1:A2)"]]);
const xlsxLoses = rows.find((r) => r.format === "xlsx")!.loses;
await cmds.run("file.exportXlsx");
const report = (await cmds.execute("file.report")).output as string[];
eq("the sparkline loss the table states is the one reported", [xlsxLoses.some((l) => l.includes("sparkline")), report.some((l) => l.includes("sparkline"))], [true, true]);
await cmds.run("formulas.use language=RexxJS");
await cmds.run("cell.set cell=C1 input==A1*2");
await cmds.run("file.exportXlsx");
const rx = (await cmds.execute("file.report")).output as string[];
eq("and the RexxJS policy is what the report says", [xlsxLoses.some((l) => l.includes("RexxJS")), rx.some((l) => l.startsWith("RexxJS formulas are not Excel's: their values written"))], [true, true]);

// expect: ok each format: read, written
// expect: ok ODS is read as cached values; XLS formulas it cannot write stay values
// expect: ok the message
// expect: ok the sparkline loss the table states is the one reported
// expect: ok and the RexxJS policy is what the report says
