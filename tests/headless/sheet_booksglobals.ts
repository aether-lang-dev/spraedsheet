// Extension lifetimes and no stale workbook state in globals (docs/comparisons/
// UNIVER.md action 4): the date system's epoch lives in dates.ts, one for the
// page, but belongs to a workbook; switching workbooks or closing one must
// leave the one now active counting from its own epoch. Triggers and macros
// are a workbook's own and go when it closes.
import { Workbooks } from "./sheet/books.ts";
import { eq } from "./lib/check.ts";

const books = new Workbooks();
const first = books.active;
await first.cmds.run("formulas.use language=Excel");
await first.cmds.run('cell.set cell=A1 input="=DATE(2026,10,21)"');
eq("the first workbook, 1900: the serial", first.wb.sheet.value("A1"), 46316);

await first.cmds.run("book.new");
const second = books.active;
await second.cmds.run("formulas.use language=Excel");
await second.cmds.run("workbook.dateSystem system=1904");
await second.cmds.run('cell.set cell=A1 input="=DATE(2026,10,21)"');
eq("the second, 1904: the same date, another serial", second.wb.sheet.value("A1"), 44854);

await second.cmds.run("book.activate name=Untitled");
await first.cmds.run("formulas.recalculate");
eq("back in the first: it counts from 1900 again, not the 1904 left in the page", [first.wb.dateSystem, first.wb.sheet.value("A1")], [1900, 46316]);
await first.cmds.run('cell.set cell=A2 input="=DATE(2026,10,21)"');
eq("a formula entered there is a 1900 serial", first.wb.sheet.value("A2"), 46316);

// And the other way: closing the one in 1900 leaves the 1904 one counting right.
await first.cmds.run("book.activate name=Book2");
await second.cmds.run('cell.set cell=A2 input="=DATE(2026,10,21)"');
eq("switched to the second: its own epoch", second.wb.sheet.value("A2"), 44854);
await second.cmds.run("book.close name=Untitled");
await second.cmds.run('cell.set cell=A3 input="=DATE(2026,10,21)"');
eq("the first closed: the second still counts from 1904", [books.all.length, second.wb.sheet.value("A3")], [1, 44854]);

// A workbook's macros and triggers go with it.
await second.cmds.run("book.new");
const third = books.active;
await third.cmds.run("macro.record name=Mark");
await third.cmds.run("cell.set cell=C1 input=marked");
await third.cmds.run("macro.stop");
await third.cmds.run("macro.trigger name=Mark on=open");
eq("the third has a macro and a trigger", [third.wb.macros.has("Mark"), third.wb.triggers.length], [true, 1]);
await third.cmds.run("book.close name=Book3");
eq("closed: the other workbook has neither, and the page's active one is the second", [second.wb.macros.size, second.wb.triggers.length, books.active === second], [0, 0, true]);
await second.cmds.run("file.new");
eq("a workbook emptied by file.new counts from 1900 again", [second.wb.dateSystem, second.wb.macros.size], [1900, 0]);
await second.cmds.run('cell.set cell=A1 input="=DATE(2026,10,21)"');
eq("and its dates follow", second.wb.sheet.value("A1"), 46316);

// expect: ok the first workbook, 1900: the serial
// expect: ok the second, 1904: the same date, another serial
// expect: ok back in the first: it counts from 1900 again, not the 1904 left in the page
// expect: ok a formula entered there is a 1900 serial
// expect: ok switched to the second: its own epoch
// expect: ok the first closed: the second still counts from 1904
// expect: ok the third has a macro and a trigger
// expect: ok closed: the other workbook has neither, and the page's active one is the second
// expect: ok a workbook emptied by file.new counts from 1900 again
// expect: ok and its dates follow
