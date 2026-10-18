// A cell's look past colour and bold, now that the datagrid draws it
// (aether-ui 08e3193d): italic, underline, strikethrough, a size, borders,
// wrapped text (formats.ts's lookStyle, and the dialects' own words for
// them); the format.font, format.border and format.wrap commands; a
// comment's corner marker and tooltip; data bars as a conditional rule.
import { lookStyle, toStyle } from "./sheet/formats.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { readSheetFile, writeSheetFile } from "./sheet/file.ts";
import { writeXlsx } from "./sheet/xlsxwrite.ts";
import { readXlsx } from "./sheet/xlsx.ts";
import { eq } from "./lib/check.ts";

eq("CSS objects (a RexxJS style function's): italic, underline, a size", toStyle({ fontStyle: "italic", textDecoration: "underline", fontSize: "16px" }),
  "font-style: italic; text-decoration: underline; font-size: 16px");
eq("the app's format strings", [toStyle("italic;underline;size:14"), toStyle("strikethrough"), toStyle("wrap")],
  ["font-style: italic; text-decoration: underline; font-size: 14px", "text-decoration: line-through", "white-space: normal"]);
eq("borders, their colours named or not", toStyle("border-bottom: 2px solid red; border: 1px solid #333"), "border-bottom: 2px solid #ff0000; border: 1px solid #333333");
eq("a look as the grid's style", lookStyle({ bold: true, italic: true, underline: true, strike: true, size: 14, wrap: true, borders: { top: "#000000", bottom: "#ff0000" } }),
  "font-weight: bold; font-style: italic; text-decoration: underline line-through; font-size: 14px; white-space: normal; border-top: 1px solid #000000; border-bottom: 1px solid #ff0000");

const wb = new Workbook();
const cmds = new Commands(wb);
const message = async (line: string) => (await cmds.run(line)).message;
await cmds.run("cell.set cell=A1 input=hello");
eq("format.font: on, then off, a part at a time", [await message("format.font range=A1:B1 italic=1 underline=1 size=14"), wb.styleOf("A1"), await message("format.font range=A1 italic=0"), wb.styleOf("A1")],
  ["A1:B1: italic, underline, size 14", "font-style: italic; text-decoration: underline; font-size: 14px", "A1: not italic", "text-decoration: underline; font-size: 14px"]);
eq("format.border: the outline of a range, or every edge, or none", [await message("format.border range=A1:B2 edge=outline"), wb.styleOf("B2").includes("border-right: 1px solid #000000"), wb.styleOf("B2").includes("border-bottom"),
  wb.styleOf("B2").includes("border-top"), await message("format.border range=A1:B2 edge=none"), wb.styleOf("B2")],
  ["A1:B2: outline borders", true, true, false, "A1:B2: no borders", ""]);
eq("an edge in a colour", [await message("format.border range=C1 edge=bottom color=#ff0000"), wb.styleOf("C1")], ["C1: bottom border", "border-bottom: 1px solid #ff0000"]);
eq("format.wrap", [await message("format.wrap range=A1"), wb.styleOf("A1").includes("white-space: normal"), wb.wrapped()], ["A1: wrapped", true, [0]]);
eq("refusals", [await message("format.border range=A1 edge=diagonal"), await message("format.font range=A1 size=x")],
  ["diagonal is not an edge (outline, all, inside, horizontal, vertical, top, right, bottom, left, none)", "x is not a size"]);
await cmds.run("edit.copy range=A1");
await cmds.run("edit.pasteSpecial at=D4 what=formats");
eq("Paste Special formats takes the look", wb.styleOf("D4"), wb.styleOf("A1"));
await cmds.run("range.clearFormats range=D4");
eq("Clear Formats clears it", wb.styleOf("D4"), "");
eq("a file keeps it", readSheetFile(JSON.parse(JSON.stringify(writeSheetFile(wb)))).sheets[0].notes.get("A1")?.look, wb.notes.get("A1").look);

await cmds.run("cell.comment cell=B3 text=checked");
eq("a comment: a corner marker, and its text to show on hover", [wb.styleOf("B3"), wb.tooltipOf("B3"), wb.tooltipOf("B4")], ["--corner: #d00000", "checked", ""]);

for (let r = 1; r <= 3; r++) await cmds.run(`cell.set cell=E${r} input=${[10, -5, 20][r - 1]}`);
eq("data bars: each a share of the largest, from the end when negative", [await message("format.dataBar range=E1:E3"), wb.styleOf("E1"), wb.styleOf("E2"), wb.styleOf("E3")],
  ["E1:E3: data bars #638ec6", "text-align: end; --fill: 0.5; --fill-color: #638ec6", "text-align: end; --fill: -0.25; --fill-color: #638ec6", "text-align: end; --fill: 1; --fill-color: #638ec6"]);
// 4.2: the font's family, a double underline, super- and subscript; the
// text's place in the cell's height, an indent, a rotation.
const wb2 = new Workbook();
const c2 = new Commands(wb2);
const m2 = async (line: string) => (await c2.run(line)).message;
await c2.run("cell.set cell=B2 input=x");
eq("format.font: a family, a double underline, a script", [await m2("format.font range=B2 family=Georgia underline=double script=super"), wb2.styleOf("B2")],
  ["B2: double underline, font Georgia, superscript", "text-decoration: underline; text-decoration-style: double; font-family: Georgia; vertical-align: super"]);
eq("and off again", [await m2('format.font range=B2 family="" underline=0 script=none'), wb2.styleOf("B2")], ["B2: not underline, the theme's font, no script", ""]);
eq("format.align: vertical, indent, rotation, alone or with align", [await m2("format.align range=B2 vertical=top indent=2 rotation=45"), wb2.styleOf("B2"),
  await m2("format.align range=B2 align=right rotation=-30"), wb2.styleOf("B2")],
  ["B2: vertically top, indent 2, turned 45°", "text-align: end; vertical-align: top; padding-left: 20px; --rotate: 45deg".replace("text-align: end; ", ""),
   "B2: aligned right, turned -30°", "text-align: end; vertical-align: top; padding-left: 20px; --rotate: -30deg"]);
eq("refusals", [await m2("format.font range=B2 script=up"), await m2("format.align range=B2 vertical=sideways"), await m2("format.align range=B2 indent=16"),
  await m2("format.align range=B2 rotation=120"), await m2("format.align range=B2"), await m2("format.font range=B2 family=<x>")],
  ["up is not a script (super, sub or none)", "sideways is not a vertical alignment (top, middle, bottom, general)", "16 is not an indent (0 to 15)",
   "120 is not a rotation (-90 to 90 degrees)", "format.align needs align=, vertical=, indent= or rotation=", "<x> is not a font's name"]);
const said2: string[] = [];
eq("SETFONT and SETALIGNMENT's third argument from REXX", [(await c2.script('SAY SETFONT("C3", "family=Verdana;size=14;underline=double;script=sub")\nSAY SETALIGNMENT("C3", "", "vertical=bottom;indent=1;rotation=90")', (l) => said2.push(l))).message,
  said2, wb2.styleOf("C3")],
  ["ran", ["C3: double underline, size 14, font Verdana, subscript", "OK"], "text-decoration: underline; font-size: 14px; text-decoration-style: double; font-family: Verdana; vertical-align: sub; vertical-align: bottom; padding-left: 10px; --rotate: 90deg"]);
eq("a cell style by example says these parts", await m2("style.define name=Turned from=C3"),
  "cell style Turned: double underline, size 14, font Verdana, subscript, bottom of the cell, indent 1, turned 90°");
const file2 = readSheetFile(JSON.parse(JSON.stringify(writeSheetFile(wb2))));
eq("a file keeps it", file2.sheets[0].notes.get("C3")?.look, { underline: "double", size: 14, family: "Verdana", script: "sub", valign: "bottom", indent: 1, rotation: 90 });
const x2 = readXlsx(writeXlsx(wb2).bytes);
const back = readSheetFile(x2.file);
eq("and XLSX: the font's name, its double underline and script, the alignment's vertical, indent and textRotation",
  [back.sheets[0].notes.get("C3")?.look, back.sheets[0].notes.get("B2")?.look], [{ underline: "double", script: "sub", family: "Verdana", size: 14, valign: "bottom", indent: 1, rotation: 90 },
  { align: "end", valign: "top", indent: 2, rotation: -30 }]);

// expect: ok CSS objects (a RexxJS style function's): italic, underline, a size
// expect: ok the app's format strings
// expect: ok borders, their colours named or not
// expect: ok a look as the grid's style
// expect: ok format.font: on, then off, a part at a time
// expect: ok format.border: the outline of a range, or every edge, or none
// expect: ok an edge in a colour
// expect: ok format.wrap
// expect: ok refusals
// expect: ok Paste Special formats takes the look
// expect: ok Clear Formats clears it
// expect: ok a file keeps it
// expect: ok a comment: a corner marker, and its text to show on hover
// expect: ok data bars: each a share of the largest, from the end when negative
// expect: ok format.font: a family, a double underline, a script
// expect: ok and off again
// expect: ok format.align: vertical, indent, rotation, alone or with align
// expect: ok refusals
// expect: ok SETFONT and SETALIGNMENT's third argument from REXX
// expect: ok a cell style by example says these parts
// expect: ok a file keeps it
// expect: ok and XLSX: the font's name, its double underline and script, the alignment's vertical, indent and textRotation
