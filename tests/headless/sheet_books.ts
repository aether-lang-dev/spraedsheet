// Several workbooks open at once (3.10): a new one beside the first, each
// with its own sheets and cells; a saved file and a sample opened beside;
// switching (Switch Windows) and closing; a name taken; the last kept
// open; one workbook alone refusing; a script that opens one going on in
// it; and the REXX functions.
import { Workbooks } from "./sheet/books.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

// A file saved before (Save names a workbook after its file, as Excel's
// does, so it is saved from a workbook of its own over the same store).
const store = memoryStore(), saver = new Commands(new Workbook(), store);
await saver.run("cell.set cell=A1 input=first");
await saver.run("file.save name=Kept");
const books = new Workbooks(store);
const switched: string[] = [];
books.onSwitch = (b) => switched.push(b.wb.name);
const first = books.active;
await first.cmds.run("cell.set cell=A1 input=first");
const n = await first.cmds.run("book.new");
eq("a new workbook beside the first, made active", [n.message, books.active.wb.name, books.active.wb.sheet.keys(), switched], ["Book2 opened (2 workbooks)", "Book2", [], ["Book2"]]);
await books.active.cmds.run("cell.set cell=A1 input=second");
eq("each its own cells", [show(first.wb.sheet.value("A1")), show(books.active.wb.sheet.value("A1"))], ["first", "second"]);
const o = await books.active.cmds.run("book.open name=Kept");
eq("a saved file opened beside them (the store shared)", [o.ok, books.active.wb.name, show(books.active.wb.sheet.value("A1")), books.all.length], [true, "Kept", "first", 3]);
const s = await books.active.cmds.run("book.open sample=budget");
const sample = books.active.wb.name;
eq("a sample", [s.ok, books.all.length, books.active.wb.sheet.keys().length > 5], [true, 4, true]);
eq("listed, the active one said", [(await first.cmds.run("book.list")).output, (await first.cmds.run("book.list")).message.endsWith("(active)")], [["Untitled", "Book2", "Kept", sample], true]);
const a = await first.cmds.run("book.activate name=book2");
eq("Switch Windows, by name in any case", [a.message, books.active.wb.name, show(books.active.wb.sheet.value("A1"))], ["Book2 active", "Book2", "second"]);
eq("closed: the one after it active, or the one before", [(await first.cmds.run("book.close")).message, (await first.cmds.execute("book.close", { name: sample })).message, books.all.map((b) => b.wb.name)],
  ["Book2 closed; Kept active", `${sample} closed; Kept active`, ["Untitled", "Kept"]]);
eq("refusals", [(await first.cmds.run("book.new name=kept")).message, (await first.cmds.run("book.activate name=Nope")).message, (await first.cmds.run("book.open")).message, (await new Commands(new Workbook()).run("book.list")).message],
  ["kept is open already", "no workbook Nope is open (Untitled, Kept)", "book.open needs a name (a saved file) or a sample", "one workbook here: no others to open beside it"]);
await first.cmds.run("book.close name=Kept");
eq("the last stays open", [(await first.cmds.run("book.close")).message, books.all.length], ["the last workbook stays open (file.new empties it)", 1]);
const said: string[] = [];
await first.cmds.script('SAY NEWWORKBOOK("Scratch")\nLET r = SETCELL("B2", 9)\nSAY GETCELL("B2")\nl = LISTWORKBOOKS()\nSAY l.0 l.2\nSAY ACTIVATEWORKBOOK("Untitled")\nSAY "[" || GETCELL("B2") || "]"\nSAY CLOSEWORKBOOK("Scratch")', (l) => said.push(l));
eq("from REXX: a script goes on in the workbook it opens", said, ["Scratch", "9", "2 Scratch", "Untitled", "[]", "Untitled"]);

// expect: ok a new workbook beside the first, made active
// expect: ok each its own cells
// expect: ok a saved file opened beside them (the store shared)
// expect: ok a sample
// expect: ok listed, the active one said
// expect: ok Switch Windows, by name in any case
// expect: ok closed: the one after it active, or the one before
// expect: ok refusals
// expect: ok the last stays open
// expect: ok from REXX: a script goes on in the workbook it opens
