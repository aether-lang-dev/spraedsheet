// Editing scenarios from IronCalc's user-model tests (docs/comparisons/IRONCALC.md
// action 1; third_party/ironcalc/PROVENANCE.md): what a copy, a paste and a cut
// carry (formats, comments, validation), merged cells through paste and cut, and
// the same through undo and redo. Translated to this app's commands; Excel's
// behaviour decides where the two differ (a merge a paste does not match is
// refused, as Excel refuses it, where IronCalc replaces it).
// Upstream: base/src/test/user_model/test_cut_n_paste.rs, test_paste_merged_cells.rs,
// test_on_paste_styles.rs, test_defined_names.rs (cut and paste).
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

async function book(lines: string[]): Promise<[Workbook, Commands]> {
  const wb = new Workbook(), cmds = new Commands(wb);
  for (const l of ["formulas.use language=Excel", ...lines]) {
    const o = await cmds.run(l);
    if (!o.ok) print(`FAIL setup ${l}: ${o.message}`);
  }
  return [wb, cmds];
}
const merges = (wb: Workbook) => wb.page(wb.active)!.merges.map((m) => `${m[0].col},${m[0].row}-${m[1].col},${m[1].row}`).sort().join(" ");

// test_on_paste_styles: a copy carries the cell's look, number format, comment and validation
{
  const [wb, cmds] = await book(["cell.set cell=A1 input=5", "format.fill range=A1 color=#ff0000", "format.font range=A1 bold=1", "format.number range=A1 format=0.00",
    "cell.comment cell=A1 text=hi", "validation.set range=A1 kind=whole min=0", "cell.set cell=C3 input=7", "format.fill range=C3 color=#00ff00"]);
  await cmds.run("edit.copy range=A1");
  await cmds.run("edit.paste at=B2");
  const a = wb.styleOf("A1");
  eq("paste carries the look and the number format", [wb.styleOf("B2") === a, wb.notes.get("B2").look?.number, wb.sheet.input("B2")], [true, "0.00", "5"]);
  eq("and the comment and the validation", [wb.notes.get("B2").comment, wb.notes.get("B2").validation !== undefined], ["hi", true]);
  await cmds.run("edit.undo");
  eq("undo takes the look, comment and validation off the target", [wb.styleOf("B2"), wb.notes.get("B2")], [wb.styleOf("Z9"), {}]);
  await cmds.run("edit.redo");
  eq("redo gives them back", [wb.styleOf("B2") === a, wb.notes.get("B2").comment], [true, "hi"]);
  // a plain, empty source replaces the target's look (Excel pastes the blank cell's formats too)
  await cmds.run("edit.copy range=E5");
  await cmds.run("edit.paste at=C3");
  eq("an empty cell copied over a coloured one leaves it plain", [wb.sheet.input("C3"), wb.styleOf("C3") === wb.styleOf("Z9")], ["", true]);
  // formulas still move, and a conditional style moves with its cell
  await cmds.run("cell.set cell=D1 input==A1*2");
  await cmds.run("edit.copy range=D1");
  await cmds.run("edit.paste at=E4");
  eq("a pasted formula moves", wb.sheet.input("E4"), "=B4*2");
}

// test_paste_merged_cells
{
  // a merged block pasted onto a plain place makes the merge there
  let [wb, cmds] = await book(["cell.set cell=B2 input=5", "cells.merge range=B2:C3"]);
  await cmds.run("edit.copy range=B2:C3");
  await cmds.run("edit.paste at=E5");
  eq("a merged block pasted makes its merge there", [merges(wb), wb.sheet.input("E5")], ["1,1-2,2 4,4-5,5", "5"]);
  await cmds.run("edit.undo");
  eq("undo takes the new merge away", merges(wb), "1,1-2,2");
  await cmds.run("edit.redo");
  eq("redo gives it back", merges(wb), "1,1-2,2 4,4-5,5");

  // paste_a_merged_block_onto_an_identical_merge
  [wb, cmds] = await book(["cell.set cell=B2 input=5", "cells.merge range=B2:C3", "cell.set cell=E2 input=9", "cells.merge range=E2:F3"]);
  await cmds.run("edit.copy range=B2:C3");
  const onto = await cmds.run("edit.paste at=E2");
  eq("a merged block onto an identical merge: it takes the value, the merges stay two", [onto.ok, merges(wb), wb.sheet.input("E2")], [true, "1,1-2,2 4,1-5,2", "5"]);
  await cmds.run("edit.undo");
  eq("undo: the target merge and its 9", [merges(wb), wb.sheet.input("E2")], ["1,1-2,2 4,1-5,2", "9"]);

  // paste_a_merged_block_onto_itself_is_a_no_op_paste
  await cmds.run("edit.copy range=B2:C3");
  await cmds.run("edit.paste at=B2");
  eq("a merged block pasted onto itself changes nothing", [merges(wb), wb.sheet.input("B2")], ["1,1-2,2 4,1-5,2", "5"]);

  // a plain block onto a merge: refused, as Excel refuses it, and nothing changed
  [wb, cmds] = await book(["cell.set cell=A1 input=1", "cell.set cell=B1 input=2", "cell.set cell=A2 input=3", "cell.set cell=B2 input=4", "cell.set cell=D1 input=x", "cells.merge range=D1:D2"]);
  await cmds.run("edit.copy range=A1:B2");
  let o = await cmds.run("edit.paste at=D1");
  eq("a plain block onto a merge it does not match: refused, nothing changed", [o.ok, o.message, merges(wb), wb.sheet.input("D1"), wb.sheet.input("E1")], [false, "To do this, all the merged cells need to be the same size.", "3,0-3,1", "x", ""]);
  // paste_partially_overlapping_a_merge_is_rejected
  [wb, cmds] = await book(["cell.set cell=A1 input=1", "cell.set cell=B1 input=2", "cells.merge range=E1:E2"]);
  await cmds.run("edit.copy range=A1:B1");
  o = await cmds.run("edit.paste at=D1");
  eq("a block pasted half over a merge: refused", [o.ok, merges(wb), wb.sheet.input("D1")], [false, "4,0-4,1", ""]);

  // paste_single_cell_into_a_merged_cell_keeps_the_merge
  [wb, cmds] = await book(["cell.set cell=A1 input=42", "cell.set cell=B2 input=x", "cells.merge range=B2:C3"]);
  await cmds.run("edit.copy range=A1");
  await cmds.run("edit.paste at=B2");
  eq("one cell pasted into a merge's anchor: the merge stays and has the value", [merges(wb), wb.sheet.input("B2")], ["1,1-2,2", "42"]);
  await cmds.run("edit.undo");
  eq("undo: x again", [merges(wb), wb.sheet.input("B2")], ["1,1-2,2", "x"]);
  await cmds.run("edit.paste at=C3");
  eq("one cell pasted onto a covered cell goes to the anchor, not under the merge", [merges(wb), wb.sheet.input("B2"), wb.sheet.input("C3")], ["1,1-2,2", "42", ""]);

  // paste_overlapping_the_source_recreates_all_merges
  [wb, cmds] = await book(["cell.set cell=A1 input=a", "cells.merge range=A1:A2", "cell.set cell=A3 input=b", "cells.merge range=A3:A4"]);
  await cmds.run("edit.copy range=A1:A4");
  await cmds.run("edit.paste at=A3");
  eq("a paste over its own source recreates every merge", [merges(wb), wb.sheet.input("A1"), wb.sheet.input("A3"), wb.sheet.input("A5")], ["0,0-0,1 0,2-0,3 0,4-0,5", "a", "a", "b"]);
  await cmds.run("edit.undo");
  eq("undo: the first two merges and b", [merges(wb), wb.sheet.input("A3"), wb.sheet.input("A5")], ["0,0-0,1 0,2-0,3", "b", ""]);
}

// test_paste_merged_cells (cut) and test_cut_n_paste
{
  // cut_paste_replaces_a_contained_target_merge: the identical target merge is the moved one
  let [wb, cmds] = await book(["cell.set cell=B2 input=5", "cells.merge range=B2:C3", "cell.set cell=E2 input=9", "cells.merge range=E2:F3"]);
  await cmds.run("edit.cut range=B2:C3");
  await cmds.run("edit.paste at=E2");
  eq("a merge cut onto an identical one: one merge there, the source gone", [merges(wb), wb.sheet.input("E2"), wb.sheet.input("B2")], ["4,1-5,2", "5", ""]);
  await cmds.run("edit.undo");
  eq("undo: both merges and both values", [merges(wb), wb.sheet.input("B2"), wb.sheet.input("E2")], ["1,1-2,2 4,1-5,2", "5", "9"]);
  await cmds.run("edit.redo");
  eq("redo", [merges(wb), wb.sheet.input("E2")], ["4,1-5,2", "5"]);

  // a merge cut onto a different-shaped merge is refused
  [wb, cmds] = await book(["cell.set cell=B2 input=5", "cells.merge range=B2:C3", "cell.set cell=E2 input=9", "cells.merge range=E2:E3"]);
  await cmds.run("edit.cut range=B2:C3");
  const o = await cmds.run("edit.paste at=E2");
  eq("a merge cut onto one of another shape: refused, nothing moved", [o.ok, merges(wb), wb.sheet.input("B2"), wb.sheet.input("E2")], [false, "1,1-2,2 4,1-4,2", "5", "9"]);

  // cut_paste_single_cell_into_a_merged_cell_keeps_the_merge
  [wb, cmds] = await book(["cell.set cell=A1 input=42", "cells.merge range=B2:C3"]);
  await cmds.run("edit.cut range=A1");
  await cmds.run("edit.paste at=B2");
  eq("one cell cut into a merge: the merge stays with the value, the source is empty", [merges(wb), wb.sheet.input("B2"), wb.sheet.input("A1")], ["1,1-2,2", "42", ""]);
  await cmds.run("edit.undo");
  eq("undo: back in A1", [merges(wb), wb.sheet.input("B2"), wb.sheet.input("A1")], ["1,1-2,2", "", "42"]);
  await cmds.run("edit.cut range=A1");
  await cmds.run("edit.paste at=C3");
  eq("cut onto a covered cell goes to the anchor", [wb.sheet.input("B2"), wb.sheet.input("C3")], ["42", ""]);
}

// test_cut_n_paste: styles move, and the target's go on undo; absolute references and names follow
{
  const [wb, cmds] = await book(["cell.set cell=A1 input=10", "format.fill range=A1 color=#ff0000", "cell.set cell=B1 input==$A$1+1", "cell.set cell=C1 input==SUM($A$1:$A$3)",
    "names.define name=TEN range=A1", "names.define name=COL range=A1:A3", "cell.set cell=A2 input=2", "cell.set cell=A3 input=3"]);
  await cmds.run("edit.cut range=A1");
  await cmds.run("edit.paste at=E5");
  eq("cut moves the cell, its look, and an absolute reference to it", [wb.sheet.input("E5"), wb.styleOf("E5").includes("#ff0000"), wb.styleOf("A1").includes("#ff0000"), wb.sheet.input("B1")], ["10", true, false, "=$E$5+1"]);
  eq("a name of the cell follows it; a name of a range it was only part of stays", [wb.names().map(([n, r]) => `${n}=${r}`)], [["COL=Sheet1!A1:A3", "TEN=Sheet1!E5"]]);
  eq("a range reaching it from outside is not changed", wb.sheet.input("C1"), "=SUM($A$1:$A$3)");
  await cmds.run("edit.undo");
  eq("undo: the look off the target, back on the source, the name and the reference back", [wb.styleOf("E5").includes("#ff0000"), wb.styleOf("A1").includes("#ff0000"), wb.sheet.input("B1"), wb.names().map(([n, r]) => `${n}=${r}`)], [false, true, "=$A$1+1", ["COL=Sheet1!A1:A3", "TEN=Sheet1!A1"]]);
  await cmds.run("edit.cut range=A1:A3");
  await cmds.run("edit.paste at=G1");
  eq("cut a whole range: an absolute range reading it, and a name of it, follow", [wb.sheet.input("C1"), wb.names().map(([n, r]) => `${n}=${r}`)], ["=SUM($G$1:$G$3)", ["COL=Sheet1!G1:G3", "TEN=Sheet1!G1"]]);
}

// test_autofill_rows `styles`: a fill carries each source cell's look, and undo takes it all back
{
  const [wb, cmds] = await book(["cell.set cell=A1 input=Alpher", "cell.set cell=A2 input=Bethe", "cell.set cell=A3 input=Gamow", "format.font range=A2 italic=1", "format.fill range=A3 color=#334455"]);
  await cmds.run("edit.fillDown range=A1:A3 source=3");
  eq("nothing to fill: the source is the whole range", wb.sheet.input("A4"), "");
  await cmds.run("edit.fillDown range=A1:A9 source=3");
  eq("fill down repeats three cells, and each look with its cell", [wb.sheet.input("A4"), wb.sheet.input("A6"), wb.styleOf("A5").includes("italic"), wb.styleOf("A6").includes("#334455"), wb.styleOf("A4").includes("#334455"), wb.styleOf("A8").includes("italic")], ["Alpher", "Gamow", true, true, false, true]);
  await cmds.run("edit.undo");
  eq("undo: the cells and the looks are gone", [wb.sheet.input("A4"), wb.styleOf("A5").includes("italic"), wb.styleOf("A6").includes("#334455")], ["", false, false]);
  await cmds.run("edit.redo");
  eq("redo: back", [wb.sheet.input("A4"), wb.styleOf("A6").includes("#334455")], ["Alpher", true]);
  await cmds.run("cell.set cell=C1 input=1");
  await cmds.run("cell.set cell=C2 input=2");
  await cmds.run("format.fill range=C1:C2 color=#00ff00");
  await cmds.run("edit.fillSeries range=C1:C5");
  eq("a series carries the look of the cells it continues", [wb.sheet.input("C5"), wb.styleOf("C3").includes("#00ff00"), wb.styleOf("C5").includes("#00ff00")], ["5", true, true]);
}

// expect: ok paste carries the look and the number format
// expect: ok and the comment and the validation
// expect: ok undo takes the look, comment and validation off the target
// expect: ok redo gives them back
// expect: ok an empty cell copied over a coloured one leaves it plain
// expect: ok a pasted formula moves
// expect: ok a merged block pasted makes its merge there
// expect: ok undo takes the new merge away
// expect: ok redo gives it back
// expect: ok a merged block onto an identical merge: it takes the value, the merges stay two
// expect: ok undo: the target merge and its 9
// expect: ok a merged block pasted onto itself changes nothing
// expect: ok a plain block onto a merge it does not match: refused, nothing changed
// expect: ok a block pasted half over a merge: refused
// expect: ok one cell pasted into a merge's anchor: the merge stays and has the value
// expect: ok undo: x again
// expect: ok one cell pasted onto a covered cell goes to the anchor, not under the merge
// expect: ok a paste over its own source recreates every merge
// expect: ok undo: the first two merges and b
// expect: ok a merge cut onto an identical one: one merge there, the source gone
// expect: ok undo: both merges and both values
// expect: ok redo
// expect: ok a merge cut onto one of another shape: refused, nothing moved
// expect: ok one cell cut into a merge: the merge stays with the value, the source is empty
// expect: ok undo: back in A1
// expect: ok cut onto a covered cell goes to the anchor
// expect: ok cut moves the cell, its look, and an absolute reference to it
// expect: ok a name of the cell follows it; a name of a range it was only part of stays
// expect: ok a range reaching it from outside is not changed
// expect: ok undo: the look off the target, back on the source, the name and the reference back
// expect: ok cut a whole range: an absolute range reading it, and a name of it, follow
// expect: ok nothing to fill: the source is the whole range
// expect: ok fill down repeats three cells, and each look with its cell
// expect: ok undo: the cells and the looks are gone
// expect: ok redo: back
// expect: ok a series carries the look of the cells it continues
