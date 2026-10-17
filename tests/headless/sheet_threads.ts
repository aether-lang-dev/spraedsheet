// Threaded comments (6.12): a thread begun at a cell, replies, Resolve and
// Reopen, a reply or the thread deleted, the comments pane's list, notes
// converted to comments, the user name they are signed with, the tooltip,
// undo, the sheet file and XLSX round trips, and the REXX functions.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("comment.author name=Ann");
const n = await cmds.run('comment.new cell=B2 text="Is this right?"');
eq("a thread begun, signed by the user name", [n.ok, n.message, wb.notes.get("B2").thread], [true, "B2: Ann: Is this right?", { author: "Ann", text: "Is this right?", replies: [] }]);
const r = await cmds.run('comment.reply cell=B2 text="Yes, checked" author=Bo');
eq("a reply, by another author", [r.message, wb.notes.get("B2").thread?.replies], ["B2: Ann: Is this right? · Bo: Yes, checked", [{ author: "Bo", text: "Yes, checked" }]]);
const twice = await cmds.run('comment.new cell=B2 text="again"');
eq("one thread a cell", [twice.ok, twice.message], [false, "B2 has a comment thread already: reply to it"]);
const res = await cmds.run("comment.resolve cell=B2");
eq("resolved, its corner grey", [res.message, wb.notes.get("B2").thread?.resolved, wb.styleOf("B2").includes("--corner: #a6a6a6")], ["B2: resolved", true, true]);
await cmds.run('comment.reply cell=B2 text="One more thing"');
eq("a reply reopens it, its corner purple again", [wb.notes.get("B2").thread?.resolved, wb.styleOf("B2").includes("--corner: #7030a0")], [undefined, true]);
await cmds.run("comment.resolve cell=B2");
eq("Reopen Thread", [(await cmds.run("comment.resolve cell=B2 off=1")).message, wb.notes.get("B2").thread?.resolved], ["B2: reopened", undefined]);
eq("the tooltip: the thread, reply by reply", wb.tooltipOf("B2"), "Ann: Is this right?\nBo: Yes, checked\nAnn: One more thing");
const d = await cmds.run("comment.delete cell=B2 reply=2");
eq("a reply deleted", [d.message, wb.notes.get("B2").thread?.replies.length, (await cmds.run("comment.delete cell=B2 reply=5")).message], ["B2: reply 2 deleted", 1, "B2's thread has no reply 5 (1 to 1)"]);
await cmds.run("cell.comment cell=D4 text=\"an old note\"");
await cmds.run("cell.comment cell=A1 text=\"another\"");
await cmds.run('comment.new cell=C9 text="Later"');
const list = await cmds.run("comment.list");
eq("the comments pane: every thread, in sheet order", list.output, ["B2: Ann: Is this right? · Bo: Yes, checked", "C9: Ann: Later"]);
const conv = await cmds.run("comment.convertNotes");
eq("notes converted to comments", [conv.message, wb.notes.get("D4").comment, wb.notes.get("D4").thread, wb.notes.get("A1").thread?.text], ["2 notes converted to comments", undefined, { author: "Ann", text: "an old note", replies: [] }, "another"]);
eq("Go To Special finds them, in sheet order", (await cmds.run("edit.goToSpecial what=comments")).output, ["A1", "B2", "D4", "C9"]);
eq("undo takes a conversion back", [(await cmds.run("edit.undo")).ok, wb.notes.get("D4").comment, wb.notes.get("D4").thread], [true, "an old note", undefined]);
const del = await cmds.run("comment.delete cell=C9");
eq("Delete Thread", [del.message, wb.notes.get("C9").thread], ["C9: thread deleted", undefined]);
eq("refusals", [(await cmds.run("comment.reply cell=E5 text=x")).message, (await cmds.run('comment.new cell=E5 text=" "')).message, (await cmds.run('comment.author name=" "')).message],
  ["E5 has no comment thread", "a comment says something", "a user name is not empty"]);
// Saved and opened again; written as XLSX and read back.
await cmds.run("comment.resolve cell=B2");
const again = new Workbook(), cmds2 = new Commands(again);
await cmds2.execute("file.import", { text: (await cmds.run("file.export")).output as string });
eq("saved with the file", again.notes.get("B2").thread, { author: "Ann", text: "Is this right?", replies: [{ author: "Bo", text: "Yes, checked" }], resolved: true });
const x = await cmds.run("file.exportXlsx");
const back = new Workbook(), cmds3 = new Commands(back);
await cmds3.run(`file.importXlsx name=T.xlsx base64=${x.output}`);
eq("an XLSX file's threaded comments, and its people", [back.notes.get("B2").thread, back.notes.get("B2").comment, back.notes.get("D4").comment], [{ author: "Ann", text: "Is this right?", replies: [{ author: "Bo", text: "Yes, checked" }], resolved: true }, undefined, "an old note"]);
const said: string[] = [];
await cmds.script('SAY ADDCOMMENT("F1", "Check F1")\nSAY REPLYCOMMENT("F1", "Done", "Cy")\nSAY RESOLVECOMMENT("F1")\nSAY GETCOMMENTTHREAD("F1")\nSAY DELETECOMMENT("F1")\nSAY "[" || GETCOMMENTTHREAD("F1") || "]"', (l) => said.push(l));
eq("from REXX", said, ["F1: Ann: Check F1", "F1: Ann: Check F1 · Cy: Done", "F1: resolved", "Ann: Check F1 · Cy: Done (resolved)", "F1: thread deleted", "[]"]);

// expect: ok a thread begun, signed by the user name
// expect: ok a reply, by another author
// expect: ok one thread a cell
// expect: ok resolved, its corner grey
// expect: ok a reply reopens it, its corner purple again
// expect: ok Reopen Thread
// expect: ok the tooltip: the thread, reply by reply
// expect: ok a reply deleted
// expect: ok the comments pane: every thread, in sheet order
// expect: ok notes converted to comments
// expect: ok Go To Special finds them, in sheet order
// expect: ok undo takes a conversion back
// expect: ok Delete Thread
// expect: ok refusals
// expect: ok saved with the file
// expect: ok an XLSX file's threaded comments, and its people
// expect: ok from REXX
