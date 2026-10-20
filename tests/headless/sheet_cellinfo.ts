// Everything non-default about a cell (blog-entry-todo.md B.3): cell.info and
// CELLINFO("B3"). A plain cell lists only its value; a styled sample cell
// lists exactly its non-defaults: value, formula, number format, look, rules
// that apply, validation, comment, name, table, and what it reads and what
// reads it.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
for (const line of [
  "formulas.use language=Excel",
  "cell.set cell=A1 input=Item", "cell.set cell=B1 input=2", 'cell.set cell=B2 input="=B1*21"', 'cell.set cell=C3 input="=SUM(B1:B2)"', "cell.set cell=D1 input=plain",
  "cell.comment cell=B2 text=answer", "cell.format cell=B2 format=number:0.00", "validation.set range=B2 kind=decimal min=0", "names.define name=answer range=B2",
  "format.fill range=B2 color=yellow", "format.font range=B2 bold=1", "format.border range=B2 edge=bottom color=red",
  "table.define name=Sales range=A1:B2 headers=1", "format.colorScale range=B1:B2", "cells.merge range=E1:F1",
]) await cmds.run(line);
const info = async (cell: string) => (await cmds.execute("cell.info", { cell })).output as Record<string, unknown>;

eq("a plain constant lists only its value", await info("D1"), { cell: "D1", value: "plain", input: "plain" });
eq("a cell with nothing lists nothing but its address", await info("H9"), { cell: "H9" });
const b2 = await info("B2");
eq("a styled formula cell lists exactly its non-defaults", b2, {
  cell: "B2", value: 42, shown: "42.00", formula: "=B1*21", format: "number:0.00",
  look: { fill: "#ffff00", bold: true, borders: { bottom: "#ff0000" } }, comment: "answer",
  validation: { kind: "decimal", min: 0, mode: "stop", when: "always" },
  rules: ["B1:B2 colorScale"], names: ["ANSWER"], table: "Sales A1:B2", reads: ["B1"], readBy: ["C3"],
});
eq("a merged cell says where", (await info("E1")).merged, "E1:F1");
eq("a cell that reads and is read, either side", [(await info("C3")).reads, (await info("B1")).readBy], [["B1", "B2"], ["B2", "C3"]]);

// REXX: a stem.
const said: string[] = [];
const r = await cmds.script('LET r = CELLINFO("B2")\nSAY r.FORMULA r.FORMAT r.COMMENT r.LOOK.FILL r.NAMES.0 r.NAMES.1 r.READS.1 r.READBY.1\nLET e = CELLINFO("H9")\nSAY e.CELL', (l) => said.push(l));
eq("CELLINFO from REXX is a stem", [r.message, said], ["ran", ["=B1*21 number:0.00 answer #ffff00 1 ANSWER B1 C3", "H9"]]);

// expect: ok a plain constant lists only its value
// expect: ok a cell with nothing lists nothing but its address
// expect: ok a styled formula cell lists exactly its non-defaults
// expect: ok a merged cell says where
// expect: ok a cell that reads and is read, either side
// expect: ok CELLINFO from REXX is a stem
