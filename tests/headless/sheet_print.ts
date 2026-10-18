// File > Print (5.8b): the sheet laid out as Export PDF lays it out, handed
// to the system's print dialog (the page's printer: sae's printing.pdf),
// the job named for the workbook; its options; no printer, said; Ctrl+P;
// PRINTSHEET from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { KEYS } from "./sheet/keys.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("cell.set cell=A1 input=Rent");
const asked: [string, string][] = [];
eq("no printer: said", (await cmds.run("file.print")).message, "printing needs the installed app (print)");
wb.printer = (b64, title) => { asked.push([String.fromCharCode(...fromBase64(b64).slice(0, 8)), title]); return ""; };
eq("printed: a PDF, the job named for the workbook", [(await cmds.run("file.print")).message, asked], ["1 page to the printer", [["%PDF-1.4", "Untitled"]]]);
eq("with Export PDF's options", (await cmds.run("file.print paper=letter landscape=1 range=A1:A1")).ok, true);
wb.printer = () => "the printer said no";
eq("a printer that refuses", (await cmds.run("file.print")).message, "the printer said no");
eq("Ctrl+P", KEYS.find((k) => k.combo === "Ctrl+P")?.id, "file.print");
wb.printer = (b64, title) => { asked.push([String.fromCharCode(...fromBase64(b64).slice(0, 8)), title]); return ""; };
const said: string[] = [];
await cmds.script('SAY PRINTSHEET()', (l) => said.push(l));
eq("from REXX", [said, asked.length], [["1 page to the printer"], 3]);

// expect: ok no printer: said
// expect: ok printed: a PDF, the job named for the workbook
// expect: ok with Export PDF's options
// expect: ok a printer that refuses
// expect: ok Ctrl+P
// expect: ok from REXX
