// Assertions in the sheet (blog-entry-todo.md B.6, app/sheet/checks.ts and
// the check.* commands): a formula that must be TRUE, with a message, kept in
// the workbook, run on every recalculation, saved with the file. A total that
// stops adding up fails its check, by name.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { checksLine } from "./sheet/checks.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
for (const line of ["cell.set cell=A1 input=Parts", "cell.set cell=B1 input=60", "cell.set cell=B2 input=40", 'cell.set cell=B3 input="=SUM(B1:B2)"', "cell.set cell=D3 input=100"]) await cmds.run(line);

const added = await cmds.execute("check.add", { name: "Total ties to report", formula: "=B3=D3", message: "B3 must equal the reported total in D3" });
eq("a check added", [added.ok, added.message], [true, "Total ties to report: B3=D3 (passing)"]);
eq("it passes: none failing, no line", [wb.checkResults.map((r) => r.ok), checksLine(wb.checkResults)], [[true], ""]);

// Run on every recalculation: a cell changed, and it fails, by name.
await cmds.run("cell.set cell=B2 input=41");
eq("a total that stops adding up fails its check, by name", wb.checkResults.map((r) => [r.name, r.ok, r.shown, r.message]), [["Total ties to report", false, "FALSE", "B3 must equal the reported total in D3"]]);
eq("the status line", checksLine(wb.checkResults), "1 check failing: Total ties to report");
await cmds.run("cell.set cell=B2 input=40");
eq("put right, it passes again", [wb.checkResults[0].ok, checksLine(wb.checkResults)], [true, ""]);

// Another, over another sheet's cell; an error is a failure; a number not 0 is a pass.
await cmds.execute("check.add", { name: "No errors", formula: "ISERROR(B3)=FALSE", message: "B3 is an error" });
await cmds.execute("check.add", { name: "Nonzero", formula: "B3", message: "B3 is zero" });
await cmds.run("cell.set cell=B1 input=0");
await cmds.run("cell.set cell=B2 input=0");
eq("a number that is 0 fails, a nonzero passes; several in the line", [wb.checkResults.map((r) => `${r.name}:${r.ok}`), checksLine(wb.checkResults)],
  [["Total ties to report:false", "No errors:true", "Nonzero:false"], "2 checks failing: Total ties to report, Nonzero"]);
await cmds.run("cell.set cell=B1 input=60");
await cmds.run("cell.set cell=B2 input=40");
await cmds.run('cell.set cell=B3 input="=1/0"');
eq("an error value fails, and says which", wb.checkResults.map((r) => [r.name, r.ok, r.shown]).filter((r) => !r[1]), [["Total ties to report", false, "#DIV/0!"], ["No errors", false, "FALSE"], ["Nonzero", false, "#DIV/0!"]]);
await cmds.run('cell.set cell=B3 input="=SUM(B1:B2)"');

// check.list: every check with its state; adding by a name again replaces it.
const list = (await cmds.execute("check.list")).output as { name: string; formula: string; message: string; ok: boolean }[];
eq("the list", list.map((c) => [c.name, c.formula, c.ok]), [["Total ties to report", "B3=D3", true], ["No errors", "ISERROR(B3)=FALSE", true], ["Nonzero", "B3", true]]);
await cmds.execute("check.add", { name: "nonzero", formula: "B3>0", message: "B3 must be above 0" });
eq("a name added again, in any case, replaces the check", ((await cmds.execute("check.list")).output as { name: string; formula: string }[]).map((c) => `${c.name} ${c.formula}`), ["Total ties to report B3=D3", "No errors ISERROR(B3)=FALSE", "nonzero B3>0"]);
eq("refused: no name, no formula, no message, one that is not there", [
  (await cmds.execute("check.add", { name: " ", formula: "TRUE", message: "m" })).message,
  (await cmds.execute("check.add", { name: "x", formula: " ", message: "m" })).message,
  (await cmds.execute("check.add", { name: "x", formula: "TRUE", message: " " })).message,
  (await cmds.execute("check.remove", { name: "nope" })).message,
], ["a check needs a name", "a check needs a formula that must be TRUE", "a check needs a message", "no check named nope"]);

// Saved with the file: the JSON, the text form, the program that makes it again.
const text = String((await cmds.execute("file.exportText")).output);
eq("the text form holds the checks, a line each", text.split("\n").filter((l) => l.startsWith("@checks")).length, 3);
const wb2 = new Workbook();
const cmds2 = new Commands(wb2);
await cmds2.execute("file.importText", { text });
eq("read back from the text: the same checks, run", [wb2.checkResults.map((r) => [r.name, r.ok]), String((await cmds2.execute("file.exportText")).output) === text], [[["Total ties to report", true], ["No errors", true], ["nonzero", true]], true]);
const wb3 = new Workbook();
const cmds3 = new Commands(wb3);
const r3 = await cmds3.script(String((await cmds.execute("file.script")).output), () => {});
eq("the workbook's program makes them again", [r3.message, String((await cmds3.execute("file.exportText")).output) === text], ["ran", true]);
const file = JSON.parse(String((await cmds.execute("file.export")).output));
eq("the file's JSON: the checks, with their sheet", file.checks, [{ name: "Total ties to report", formula: "B3=D3", message: "B3 must equal the reported total in D3", sheet: "Sheet1" }, { name: "No errors", formula: "ISERROR(B3)=FALSE", message: "B3 is an error", sheet: "Sheet1" }, { name: "nonzero", formula: "B3>0", message: "B3 must be above 0", sheet: "Sheet1" }]);

// A file opened is checked at once.
await cmds2.run("cell.set cell=D3 input=99");
eq("a changed workbook read back is checked from the start", wb2.checkResults.filter((r) => !r.ok).map((r) => r.name), ["Total ties to report"]);

// Another sheet: a check keeps the sheet it was added on.
await cmds.run("sheet.add name=Costs");
await cmds.run("cell.set cell=B3 input=7");
await cmds.execute("check.add", { name: "Costs seven", formula: "B3=7", message: "Costs!B3 is 7" });
await cmds.run("sheet.activate name=Sheet1");
eq("a check is worked out on the sheet it was added on, whichever is active", wb.checkResults.find((r) => r.name === "Costs seven")!.ok, true);

// Removed.
await cmds.execute("check.remove", { name: "NO ERRORS" });
eq("removed, in any case", wb.checks.map((c) => c.name), ["Total ties to report", "nonzero", "Costs seven"]);

// REXX.
const said: string[] = [];
await cmds.script('LET a = ADDCHECK("Tied", "B3=D3", "B3 is not D3")\nSAY a\nLET r = SETCELL("D3", 5)\nLET f = CHECKS()\nSAY f.0 f.1\nLET r = SETCELL("D3", 100)\nLET f = CHECKS()\nSAY f.0\nLET x = REMOVECHECK("Tied")\nSAY x', (l) => said.push(l));
eq("ADDCHECK, CHECKS and REMOVECHECK from REXX", said, ["Tied: B3=D3 (passing)", "2 Total ties to report: B3 must equal the reported total in D3", "0", "removed Tied"]);

// expect: ok a check added
// expect: ok it passes: none failing, no line
// expect: ok a total that stops adding up fails its check, by name
// expect: ok the status line
// expect: ok put right, it passes again
// expect: ok a number that is 0 fails, a nonzero passes; several in the line
// expect: ok an error value fails, and says which
// expect: ok the list
// expect: ok a name added again, in any case, replaces the check
// expect: ok refused: no name, no formula, no message, one that is not there
// expect: ok the text form holds the checks, a line each
// expect: ok read back from the text: the same checks, run
// expect: ok the workbook's program makes them again
// expect: ok the file's JSON: the checks, with their sheet
// expect: ok a changed workbook read back is checked from the start
// expect: ok a check is worked out on the sheet it was added on, whichever is active
// expect: ok removed, in any case
// expect: ok ADDCHECK, CHECKS and REMOVECHECK from REXX
