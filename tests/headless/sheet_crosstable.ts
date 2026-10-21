// A table is the workbook's, as in Excel: Sales[Amount] written on another
// sheet than the table's reads its column (docs/comparisons/LIBREOFFICE.md
// action 1's gap: it gave #REF!). It follows the table through a row inserted
// in it, a renamed column, a table deleted, undo and redo, and a reopened file.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const bad: string[] = [];
for (const line of [
  "formulas.use language=Excel",
  "cell.set cell=A1 input=Region", "cell.set cell=B1 input=Amount",
  "cell.set cell=A2 input=North", "cell.set cell=B2 input=10",
  "cell.set cell=A3 input=South", "cell.set cell=B3 input=30",
  "cell.set cell=A4 input=East", "cell.set cell=B4 input=60",
  "table.define name=Sales range=A1:B4 headers=1",
  "sheet.add name=Report",
  'cell.set cell=A1 input="=SUM(Sales[Amount])"',
  'cell.set cell=A2 input="=ROWS(Sales)"',
  'cell.set cell=A3 input="=Sales[[#Headers],[Amount]]"',
  'cell.set cell=A4 input="=COUNTA(Sales[Region])"',
  "sheet.activate name=Sheet1",
]) {
  const o = await cmds.run(line);
  if (!o.ok) bad.push(`${line}: ${o.message}`);
}
eq("set up", bad, []);
const r = () => ["A1", "A2", "A3", "A4"].map((k) => wb.page("Report")!.sheet.value(k));
eq("the report reads the table on the other sheet", r(), [100, 3, "Amount", 3]);

// A table's name alone, and its own row, refer to the table's sheet too.
await cmds.run("sheet.activate name=Report");
await cmds.run('cell.set cell=B1 input="=SUM(Sales)"');
await cmds.run('cell.set cell=B2 input="=SUM(Sales[[Amount]])*2"');
eq("a table's name alone, and a doubled sum", [wb.sheet.value("B1"), wb.sheet.value("B2")], [100, 200]);
await cmds.run("sheet.activate name=Sheet1");

// A row inserted into the table: the report's sum takes it in.
await cmds.run("rows.insert at=3");
await cmds.run("cell.set cell=A3 input=West");
await cmds.run("cell.set cell=B3 input=5");
eq("a row inserted into the table, one value in it: the report follows", r(), [105, 4, "Amount", 4]);

// A column renamed: the reference to its old name is #REF!.
await cmds.run("cell.set cell=B1 input=Total");
eq("the column renamed: Amount is no longer found, on the other sheet either", [String(wb.page("Report")!.sheet.value("A1")).startsWith("#REF!"), wb.page("Report")!.sheet.value("A3").toString().startsWith("#REF!")], [true, true]);
await cmds.run("cell.set cell=B1 input=Amount");
eq("and back", r(), [105, 4, "Amount", 4]);

// Undo and redo.
await cmds.run("edit.undo");
await cmds.run("edit.undo");
eq("undo twice (the name, then the rename back): the report reads the table as it was", r()[0], 105);
await cmds.run("edit.undo");
await cmds.run("edit.undo");
await cmds.run("edit.undo");
eq("undo the value, the new cell and the insert: the report reads the first three rows", r(), [100, 3, "Amount", 3]);
await cmds.run("edit.redo");
await cmds.run("edit.redo");
await cmds.run("edit.redo");
eq("redo them: the report reads four rows", r(), [105, 4, "Amount", 4]);

// Saved and reopened, as a file and as text.
const wb2 = new Workbook(), cmds2 = new Commands(wb2);
await cmds2.execute("file.import", { text: String((await cmds.execute("file.export")).output) });
eq("reopened from its file", [wb2.page("Report")!.sheet.value("A1"), wb2.page("Report")!.sheet.input("A1")], [105, "=SUM(Sales[Amount])"]);
const wb3 = new Workbook(), cmds3 = new Commands(wb3);
await cmds3.execute("file.importText", { text: String((await cmds.execute("file.exportText")).output) });
eq("reopened from its text form", wb3.page("Report")!.sheet.value("A1"), 105);

// A table of the same name on the sheet itself comes first; the sheet's own name for it.
await cmds.run("sheet.activate name=Report");
await cmds.run("cell.set cell=D1 input=Amount");
await cmds.run("cell.set cell=D2 input=1");
await cmds.run("cell.set cell=D3 input=2");
await cmds.run("table.define name=Mine range=D1:D3 headers=1");
await cmds.run('cell.set cell=C1 input="=SUM(Mine[Amount])"');
eq("a table of this sheet beside the other's", [wb.sheet.value("C1"), wb.sheet.value("A1")], [3, 105]);

// The table deleted: the report's reference is #REF!, and back with undo.
await cmds.run("sheet.activate name=Sheet1");
await cmds.run("table.delete name=Sales");
eq("the table deleted: the report's sum is #REF!", String(wb.page("Report")!.sheet.value("A1")).startsWith("#REF!"), true);
await cmds.run("edit.undo");
eq("undone: the report reads it again", wb.page("Report")!.sheet.value("A1"), 105);

// Convert to Range: the report's reference becomes the cells, with their sheet.
await cmds.run("table.toRange name=Sales");
eq("converted to a range: the report reads the cells, and says so", [wb.page("Report")!.sheet.input("A1"), wb.page("Report")!.sheet.input("A2"), wb.page("Report")!.sheet.value("A1")], ["=SUM(Sheet1!$B$2:$B$5)", "=ROWS(Sheet1!$A$2:$B$5)", 105]);
await cmds.run("edit.undo");
eq("undone: the structured reference is back", [wb.page("Report")!.sheet.input("A1"), wb.page("Report")!.sheet.value("A1")], ["=SUM(Sales[Amount])", 105]);

// expect: ok set up
// expect: ok the report reads the table on the other sheet
// expect: ok a table's name alone, and a doubled sum
// expect: ok a row inserted into the table, one value in it: the report follows
// expect: ok the column renamed: Amount is no longer found, on the other sheet either
// expect: ok and back
// expect: ok undo twice (the name, then the rename back): the report reads the table as it was
// expect: ok undo the value, the new cell and the insert: the report reads the first three rows
// expect: ok redo them: the report reads four rows
// expect: ok reopened from its file
// expect: ok reopened from its text form
// expect: ok a table of this sheet beside the other's
// expect: ok the table deleted: the report's sum is #REF!
// expect: ok undone: the report reads it again
// expect: ok converted to a range: the report reads the cells, and says so
// expect: ok undone: the structured reference is back
