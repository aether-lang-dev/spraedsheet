// Record Macro as REXX (7.1): a session's commands recorded as a RexxJS
// program (SETCELL for a cell, ADDRESS SHEET for the rest), which, run on
// a new workbook, makes the same workbook; not what only looks (a zoom) or
// reads; edited (a loop) and run again; saved with the workbook; one step
// of undo; refusals; and the REXX functions. And (7.2) a macro run with
// parameters, on a change to a range, on open.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeSheetFile } from "./sheet/file.ts";
import { toRexx } from "./sheet/recorder.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("the program a journal renders to", toRexx("Totals", ["cell.set cell=A1 input=5", 'cell.set cell=A2 input="=A1*2"', "format.font range=A1:A2 bold=1", 'cell.comment cell=A1 text="say ""hi"""']),
  ['/* Totals: recorded by the spreadsheet, 4 commands */', 'CALL SETCELL "A1", "5"', 'CALL SETCELL "A2", "=A1*2"', 'ADDRESS SHEET "format.font range=A1:A2 bold=1"', 'ADDRESS SHEET "cell.comment cell=A1 text=""say """"hi"""""""'].join("\n"));

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const started = await cmds.run("macro.record name=Budget");
await cmds.run("cell.set cell=A1 input=Rent");
await cmds.run("cell.set cell=B1 input=1200");
await cmds.run("cell.set cell=B2 input=300");
await cmds.run('cell.set cell=B3 input="=SUM(B1:B2)"');
await cmds.run("format.number range=B1:B3 format=currency");
await cmds.run("view.zoom percent=150");
await cmds.run("cell.get cell=B3");
await cmds.run("sheet.add name=Notes");
await cmds.run('cell.set cell=A1 input="a second sheet"');
await cmds.run("sheet.activate name=Sheet1");
const stopped = await cmds.run("macro.stop");
eq("recorded: what changed the workbook, not a zoom or a read", [started.message, stopped.message, String(stopped.output).split("\n").length, String(stopped.output).includes("view.zoom"), String(stopped.output).includes("cell.get")],
  ["recording Budget", "Budget: 8 commands recorded", 9, false, false]);
const again = new Workbook(), cmds2 = new Commands(again);
await cmds2.run("formulas.use language=Excel");
await cmds2.execute("macro.set", { name: "Budget", text: String(stopped.output) });
const ran = await cmds2.run("macro.run name=Budget");
const same = (a: Workbook, b: Workbook) => { const { macros: _a, ...x } = writeSheetFile(a) as Record<string, unknown>; const { macros: _b, ...y } = writeSheetFile(b) as Record<string, unknown>; return JSON.stringify(x) === JSON.stringify(y); };
eq("run on a new workbook, it makes the same one", [ran.message, show(again.sheet.value("B3")), again.textOf("B3"), again.pages().map((p) => p.name), same(wb, again)], ["Budget ran", "1500", "$1,500.00", ["Sheet1", "Notes"], true]);
eq("one step of undo", [(await cmds2.run("edit.undo")).ok, again.sheet.input("B3"), again.pages().length], [true, "", 1]);
// Edited: a loop, then run.
await cmds.execute("macro.set", { name: "Fill", text: 'DO i = 1 TO 3\n  CALL SETCELL "D" || i, i * 10\nEND\nSAY "filled"' });
const filled = await cmds.run("macro.run name=Fill");
eq("an edited macro: a loop", [filled.message, filled.output, show(wb.sheet.value("D3"))], ["Fill ran", ["filled"], "30"]);
eq("listed, and saved with the workbook", [(await cmds.run("macro.list")).output, Object.keys((writeSheetFile(wb) as { macros: object }).macros)], [["Budget", "Fill"], ["Budget", "Fill"]]);
const third = new Workbook(), cmds3 = new Commands(third);
await cmds3.execute("file.import", { text: JSON.stringify(writeSheetFile(wb)) });
eq("opened again, the macros with it", [(await cmds3.run("macro.list")).output, (await cmds3.run("macro.get name=Fill")).output], [["Budget", "Fill"], 'DO i = 1 TO 3\n  CALL SETCELL "D" || i, i * 10\nEND\nSAY "filled"']);
// Recording while a macro runs: the run, not what it did.
await cmds.run("macro.record name=Twice");
await cmds.run("macro.run name=Fill");
eq("a macro run while recording is one line", (await cmds.run("macro.stop")).output, '/* Twice: recorded by the spreadsheet, 1 command */\nADDRESS SHEET "macro.run name=Fill"');
eq("refusals", [(await cmds.run("macro.record name=9lives")).message, (await cmds.run("macro.stop")).message, (await cmds.run("macro.run name=Nope")).message, (await cmds.run("macro.delete name=Nope")).message],
  ["9lives is not a macro's name (a letter, then letters, digits or _)", "nothing is being recorded", "no macro Nope", "no macro Nope"]);
// 7.2: parameters, and triggers. (SETCELL called as a function where an
// argument is an upper-case variable alone: RexxJS's CALL still takes such
// a name for a function, "CALL SETCELL CELL, ..." failing; RexxJS b3d5e4e.)
await cmds.execute("macro.set", { name: "Put", text: 'r = SETCELL(CELL, N * 2)' });
eq("a macro run with parameters", [(await cmds.run('macro.run name=Put args="CELL=H1;N=21"')).message, show(wb.sheet.value("H1")), (await cmds.run('macro.run name=Put args="1x=2"')).message], ["Put ran", "42", "1x=2 is not a parameter (NAME=value)"]);
await cmds.execute("macro.set", { name: "Twice", text: 'CALL SETCELL "J1", GETCELL("I1") * 2' });
eq("on a change to a range: run after it", [(await cmds.run("macro.trigger name=Twice on=change range=I1:I5")).message, (await cmds.run("cell.set cell=I1 input=8")).ok, show(wb.sheet.value("J1"))],
  ["Twice runs after a change to Sheet1!I1:I5", true, "16"]);
await cmds.run("cell.set cell=K1 input=1");
eq("not after one elsewhere, nor from inside itself", [show(wb.sheet.value("J1")), (await cmds.run("cell.set cell=I1 input=9")).ok, show(wb.sheet.value("J1"))], ["16", true, "18"]);
await cmds.execute("macro.set", { name: "Hello", text: 'CALL SETCELL "L1", "opened"' });
await cmds.run("macro.trigger name=Hello on=open");
const opened = new Workbook(), oc = new Commands(opened);
await oc.execute("file.import", { text: JSON.stringify(writeSheetFile(wb)) });
eq("on open: run when the workbook is opened, the triggers saved with it", [show(opened.sheet.value("L1")), opened.triggers.map((t) => `${t.macro} ${t.on}`)], ["opened", ["Twice change", "Hello open"]]);
eq("taken away", [(await cmds.run("macro.trigger name=Twice on=change range=I1:I5 off=1")).message, (await cmds.run("macro.trigger name=Nope on=open")).message], ["Twice no longer runs on a change to I1:I5", "no macro Nope"]);
await cmds.run("macro.trigger name=Hello on=open off=1");
const said: string[] = [];
await cmds.script('SAY RECORDMACRO("R")\nLET r = SETCELL("F1", 7)\nm = STOPRECORDING()\nSAY m\nSAY RUNMACRO("R")\nl = LISTMACROS()\nSAY l.0 l.3\nSAY DELETEMACRO("R")', (l) => said.push(l));
eq("from REXX", said, ["recording R", '/* R: recorded by the spreadsheet, 1 command */\nCALL SETCELL "F1", "7"', "R ran", "6 Hello", "R deleted"]);
const said2: string[] = [];
await cmds.script('SAY RUNMACRO("Put", "CELL=H2;N=5")\nSAY GETCELL("H2")\nSAY ONCHANGE("I1:I5", "Twice")\nSAY ONCHANGE("I1:I5", "Twice", 1)\nSAY ONOPEN("Hello")', (l) => said2.push(l));
eq("7.2 from REXX", said2, ["Put ran", "10", "Twice runs after a change to Sheet1!I1:I5", "Twice no longer runs on a change to I1:I5", "Hello runs when the workbook opens"]);

const said3: string[] = [];
await cmds.script('LET r = CALL twice 21\nSAY r\nCALL twice 4\nSAY RESULT\nEXIT\ntwice:\nPARSE ARG n\nRETURN n * 2', (l) => said3.push(l));
eq("LET r = CALL f gives f's RETURN value, as RESULT does (RexxJS 5d96050)", said3, ["42", "8"]);
// expect: ok the program a journal renders to
// expect: ok recorded: what changed the workbook, not a zoom or a read
// expect: ok run on a new workbook, it makes the same one
// expect: ok one step of undo
// expect: ok an edited macro: a loop
// expect: ok listed, and saved with the workbook
// expect: ok opened again, the macros with it
// expect: ok a macro run while recording is one line
// expect: ok refusals
// expect: ok a macro run with parameters
// expect: ok on a change to a range: run after it
// expect: ok not after one elsewhere, nor from inside itself
// expect: ok on open: run when the workbook is opened, the triggers saved with it
// expect: ok taken away
// expect: ok from REXX
// expect: ok 7.2 from REXX
// expect: ok LET r = CALL f gives f's RETURN value, as RESULT does (RexxJS 5d96050)
