// Pictures in cells (6.11c): Place in Cell from a picture's bytes, with its
// alternative text, and a floating picture placed in the cell it sits on;
// IMAGE from a data: URL (no network); refusals; taken out; undo; saved
// with the file; and the REXX functions.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const RED = "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEUlEQVR4nGM4oaGBFTEMLQkAgl1GAWqNFmsAAAAASUVORK5CYII=";
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
eq("Place in Cell", [(await cmds.run(`insert.cellPicture cell=B2 base64=${RED} alt=logo`)).message, wb.notes.get("B2").picture?.alt, wb.pictureIn("B2") === RED], ["a picture in B2: logo", "logo", true]);
await cmds.run(`insert.picture base64=${RED} at=D4`);
eq("a floating picture placed in its cell", [(await cmds.run('insert.cellPicture name="Picture 1"')).message, wb.page("Sheet1")!.drawings.length, wb.pictureIn("D4") === RED], ["Picture 1 placed in D4", 0, true]);
await cmds.run(`cell.set cell=C3 input="=IMAGE(""data:image/png;base64,${RED}"", ""a red square"")"`);
eq("IMAGE from a data: URL: its value the alternative text, the picture drawn", [show(wb.sheet.value("C3")), wb.pictureIn("C3") === RED], ["a red square", true]);
await cmds.run('cell.set cell=C4 input="=IMAGE(""https://example.org/a.png"")"');
eq("no pictures from the network", [show(wb.sheet.value("C4")), wb.pictureIn("C4")], ["#VALUE!", ""]);
eq("refusals", [(await cmds.run("insert.cellPicture cell=E5 base64=aGVsbG8=")).message, (await cmds.run("insert.cellPicture")).message], ["not a picture (PNG, JPEG, GIF or BMP)", "Place in Cell needs cell= (and base64=) or name="]);
const again = new Workbook(), cmds2 = new Commands(again);
await cmds2.execute("file.import", { text: (await cmds.run("file.export")).output as string });
eq("saved with the file", [again.pictureIn("B2") === RED, again.notes.get("B2").picture?.alt], [true, "logo"]);
eq("taken out, and undone", [(await cmds.run("cell.removePicture range=A1:D4")).message, wb.pictureIn("B2"), (await cmds.run("edit.undo")).ok, wb.pictureIn("B2") === RED], ["2 pictures taken out of A1:D4", "", true, true]);
const said: string[] = [];
await cmds.script(`SAY INSERTCELLPICTURE("F6", "${RED}", "dot")\nSAY REMOVECELLPICTURE("F6")`, (l) => said.push(l));
eq("from REXX", said, ["a picture in F6: dot", "1 picture taken out of F6"]);

// expect: ok Place in Cell
// expect: ok a floating picture placed in its cell
// expect: ok IMAGE from a data: URL: its value the alternative text, the picture drawn
// expect: ok no pictures from the network
// expect: ok refusals
// expect: ok saved with the file
// expect: ok taken out, and undone
// expect: ok from REXX
