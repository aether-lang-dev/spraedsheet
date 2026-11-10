// FORMULATEXT and ISFORMULA read what a cell holds as typed, not its value: the cell they
// name is no cycle through them (Excel: =FORMULATEXT(A1) in A1 is its own text), and they
// are worked out again when that cell is retyped. (IronCalc's FORMULATEXT.xlsx, CHOOSE.xlsx.)
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), sh = wb.sheet;
const at = (k: string): string => show(sh.value(k));

// A formula asking about its own cell, and one asking about a cell that reads it.
await sh.set("A1", "=FORMULATEXT(A1)");
eq("FORMULATEXT of its own cell", at("A1"), "=FORMULATEXT(A1)");
await sh.set("F1", "=E1");
await sh.set("E1", "=CONCATENATE(FORMULATEXT(F1), \" is fine\")");
eq("FORMULATEXT of a cell that reads this one", [at("E1"), at("F1")], ["=E1 is fine", "=E1 is fine"]);
await sh.set("A3", "=B3");
await sh.set("B3", "=FORMULATEXT(A3)");
eq("A3 reads B3, which reads A3's formula", [at("A3"), at("B3")], ["=B3", "=B3"]);
await sh.set("A5", "=E1-6+3");
eq("what reads such a cell reads its text", at("A5"), "#VALUE!");

// A cycle of values is still one.
await sh.set("C1", "=C2");
await sh.set("C2", "=C1");
eq("a cycle of values", [at("C1"), at("C2")], ["#CYCLE!", "#CYCLE!"]);

// Retyped, what asks about the cell is worked out again.
await sh.set("D1", "5");
await sh.set("D2", "=ISFORMULA(D1)");
await sh.set("D3", "=FORMULATEXT(D1:D2)");
eq("asked of a constant", [at("D2"), at("D3")], ["FALSE", "#N/A"]);
await sh.set("D1", "=1+1");
eq("retyped as a formula", [at("D2"), at("D3")], ["TRUE", "=1+1"]);
await sh.set("D1", "=2*3");
eq("retyped again", at("D3"), "=2*3");

// expect: ok FORMULATEXT of its own cell
// expect: ok FORMULATEXT of a cell that reads this one
// expect: ok A3 reads B3, which reads A3's formula
// expect: ok what reads such a cell reads its text
// expect: ok a cycle of values
// expect: ok asked of a constant
// expect: ok retyped as a formula
// expect: ok retyped again
