// Validation policy per entry path (docs/comparisons/LIBREOFFICE.md action 3): typing in a
// cell and a script's SETCELL are refused when they break a cell's rule; a paste
// and an import bring their cells in as Excel's do, and say so when they broke
// a rule, instead of passing silently; validation.policy lists which is which.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await cmds.run("validation.set range=A1:A3 kind=whole min=1 max=10");

const typed = await cmds.run("cell.set cell=A1 input=99");
eq("typing a value that breaks the rule is refused", [typed.ok, wb.sheet.input("A1")], [false, ""]);
const said: string[] = [];
const scripted = await cmds.script('LET x = SETCELL("A1", 55)', (l) => said.push(l));
eq("and so is a script's SETCELL", [scripted.ok, wb.sheet.input("A1")], [false, ""]);

await cmds.run("cell.set cell=C1 input=77");
await cmds.run("edit.copy range=C1");
const pasted = await cmds.run("edit.pasteSpecial at=A2 what=values");
eq("a paste of values brings the value in, and says it broke the rule", [pasted.ok, wb.sheet.input("A2"), pasted.message.endsWith("; 1 cell breaks its validation (A2): validation.circle shows them")], [true, "77", true]);
const text = await cmds.run("edit.pasteText at=A3 text=88");
eq("so does pasted text", [text.ok, wb.sheet.input("A3"), text.message.endsWith("1 cell breaks its validation (A3): validation.circle shows them")], [true, "88", true]);
const csv = await cmds.run('file.importCsv at=A1 text="500"');
eq("and an import", [csv.ok, wb.sheet.input("A1"), csv.message.includes("breaks its validation (A1)")], [true, "500", true]);
const fine = await cmds.run("cell.set cell=B1 input=3");
const ok = await cmds.run("edit.copy range=B1");
const pastedOk = await cmds.run("edit.pasteSpecial at=A2 what=values");
eq("a paste that breaks nothing says nothing of it", [pastedOk.ok, pastedOk.message.includes("validation")], [true, false]);
// The message is not repeated for cells that were invalid already.
const again = await cmds.run("edit.paste at=A1");
eq("nor for the same cell once more when it was not the paste that broke it", again.message.includes("validation"), false);

// The policy, listed.
const policy = (await cmds.execute("validation.policy")).output as { path: string; commands: string; checked: boolean }[];
eq("validation.policy lists each path and whether it is checked", policy.map((p) => [p.path, p.checked]),
  [["typing in a cell", true], ["a script or the bus", true], ["a paste", false], ["an import", false], ["a fill", false]]);
eq("and the commands of each", policy[0].commands, "cell.set");

// expect: ok typing a value that breaks the rule is refused
// expect: ok and so is a script's SETCELL
// expect: ok a paste of values brings the value in, and says it broke the rule
// expect: ok so does pasted text
// expect: ok and an import
// expect: ok a paste that breaks nothing says nothing of it
// expect: ok nor for the same cell once more when it was not the paste that broke it
// expect: ok validation.policy lists each path and whether it is checked
// expect: ok and the commands of each
