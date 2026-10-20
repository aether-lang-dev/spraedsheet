// Files: the workbook saved to and opened from the page's storage by name
// (every sheet, its notes and names, in the Tauri app's format), the saved
// files listed; a new workbook; CSV out (what the selection, or everything
// used, shows; or TSV, or HTML) and in (at the selected cell, its delimiter
// detected; or JSON rows, or an HTML page's table); the Excel example, the Tauri
// app's three sample sheets (RexxJS formulas), an empty sheet, and the way
// home. As an installed app with `documents`, files the person picks too:
// a workbook in the Tauri app's format (.json, .rexxsheet), an XLSX (an
// encrypted one by the password field's password) or CSV opened,
// the workbook saved, the selection exported as CSV; the page never sees
// where they are.
import type { SheetView } from "./view.ts";
import { TEXT_HEADER } from "../textform.ts";
import { changeLine } from "../gitstore.ts";
import type { Change } from "../gitstore.ts";

export function fileBar(v: SheetView): void {
  const name = state(""), csv = state(""), saved = state(""), shown = state("");
  const listed = async () => {
    const o = await v.cmds.execute("file.list");
    saved.set(o.message);
  };
  ui.wrap(() => {
    bind(ui.textfield("file name", () => {}), name);
    v.btn("Save", async () => { await v.run("file.save", name.value.trim() ? { name: name.value } : {}); await listed(); });
    v.btn("Open", () => v.run("file.open", { name: name.value }));
    v.btn("Delete file", async () => { await v.run("file.delete", { name: name.value }); await listed(); });
    v.btn("New", () => v.run("file.new"));
    bind(ui.text(""), saved);
  });
  ui.wrap(() => {
    v.btn("Export CSV", async () => {
      const s = v.selection(), many = s && (s[0].col !== s[1].col || s[0].row !== s[1].row);
      const o = await v.run("file.exportCsv", many ? { range: v.range() } : {});
      shown.set(String(o.output ?? ""));
    });
    v.btn("Export TSV", async () => {
      const s = v.selection(), many = s && (s[0].col !== s[1].col || s[0].row !== s[1].row);
      const o = await v.run("file.exportCsv", { ...(many ? { range: v.range() } : {}), delimiter: "tab" });
      shown.set(String(o.output ?? ""));
    });
    v.btn("Export PDF", () => v.run("file.exportPdf", { gridlines: "1" }));
    // File > Print: the system's print dialog (the installed app's print).
    v.btn("Print", () => v.run("file.print", { gridlines: "1" }));
    bind(ui.textarea("CSV"), csv);
    v.btn("Import CSV", () => v.run("file.importCsv", { at: v.here(), text: csv.value }));
    v.btn("Import JSON", () => v.run("file.importJson", { at: v.here(), text: csv.value }));
    v.btn("Import HTML", () => v.run("file.importHtml", { at: v.here(), text: csv.value }));
    v.btn("Export HTML", async () => {
      const s = v.selection(), many = s && (s[0].col !== s[1].col || s[0].row !== s[1].row);
      const o = await v.run("file.exportHtml", { ...(many ? { range: v.range() } : {}), header: "1" });
      shown.set(String(o.output ?? ""));
    });
  });
  bind(ui.text_wrapped("", 720), shown);
  if (typeof documents !== "undefined") documentsRow(v);
  ui.wrap(() => {
    v.btn("Paste example", () => v.run("example.paste"));
    v.btn("Budget (RexxJS)", () => v.run("file.open", { sample: "budget" }));
    v.btn("Test sheet (RexxJS)", () => v.run("file.open", { sample: "test" }));
    v.btn("Formatted (RexxJS)", () => v.run("file.open", { sample: "formatted" }));
    v.btn("Clear all", () => v.run("sheet.clear"));
    v.btn("Home", () => browserContext.changePage("/"));
  });
  listed();
}

// Open file, Save file and Export CSV file, through the person's picker.
function documentsRow(v: SheetView): void {
  // What failed, on the status line, as a refused command's reason is.
  const tried = (work: () => Promise<void>) => async () => {
    try {
      await work();
    } catch (e) {
      v.status.set(String((e as Error)?.message ?? e));
    }
  };
  // An encrypted XLSX's password, typed here before Open file (its text
  // shown nowhere).
  let password = "";
  ui.wrap(() => {
    ui.securefield("file password", (t: string) => { password = t; });
    v.btn("Open file", tried(async () => {
      const d = await documents.pick_open("Open a workbook, an XLSX, or a CSV file");
      if (!d) return;
      // An XLSX is bytes: read as base64 (sae's documents.read_base64).
      if (/\.xlsx$/i.test(d.name)) {
        await v.run("file.importXlsx", { base64: await d.read_base64(), name: d.name, ...(password ? { password } : {}) });
        return;
      }
      if (/\.xls$/i.test(d.name)) {
        await v.run("file.importXls", { base64: await d.read_base64(), name: d.name });
        return;
      }
      if (/\.ods$/i.test(d.name)) {
        await v.run("file.importOds", { base64: await d.read_base64(), name: d.name });
        return;
      }
      const text = await d.read_text();
      // The canonical text (textform.ts) by its first line, whatever its name (.sheet.txt).
      if (text.startsWith(TEXT_HEADER)) await v.run("file.importText", { text });
      else if (/\.(csv|tsv|tab|txt)$/i.test(d.name)) await v.run("file.importCsv", { at: v.here(), text, ...(/\.(tsv|tab)$/i.test(d.name) ? { delimiter: "tab" } : {}) });
      else if (/\.html?$/i.test(d.name)) await v.run("file.importHtml", { at: v.here(), text });
      // JSON rows (an array), or a workbook in the Tauri app's format.
      else if (/^\s*\[/.test(text)) await v.run("file.importJson", { at: v.here(), text });
      else await v.run("file.import", { text });
    }));
    v.btn("Save file", tried(async () => {
      const d = await documents.pick_save(`${v.wb.name}.json`);
      if (!d) return;
      const o = await v.cmds.execute("file.export");
      await d.write_text(String(o.output));
      print(`sheet: saved ${d.name}`);
      v.status.set(`Saved ${d.name}`);
    }));
    // The canonical text (textform.ts): a cell a line, for a diff.
    v.btn("Save text file", tried(async () => {
      const d = await documents.pick_save(`${v.wb.name}.sheet.txt`);
      if (!d) return;
      const o = await v.cmds.execute("file.exportText");
      await d.write_text(String(o.output));
      print(`sheet: saved ${d.name}`);
      v.status.set(`Saved ${d.name}`);
    }));
    v.btn("Save XLSX file", tried(async () => {
      const d = await documents.pick_save(`${v.wb.name}.xlsx`);
      if (!d) return;
      const o = await v.cmds.execute("file.exportXlsx");
      await d.write_base64(String(o.output));
      print(`sheet: saved ${d.name}`);
      v.status.set(`Saved ${d.name}${v.wb.fileReport.length ? `: ${v.wb.fileReport.join("; ")}` : ""}`);
    }));
    v.btn("Export PDF file", tried(async () => {
      const d = await documents.pick_save(`${v.wb.name}.pdf`);
      if (!d) return;
      const o = await v.cmds.execute("file.exportPdf", { gridlines: "1" });
      await d.write_base64(String(o.output));
      print(`sheet: exported ${d.name}: ${o.message}`);
      v.status.set(`Exported ${d.name}`);
    }));
    v.btn("Export CSV file", tried(async () => {
      const s = v.selection(), many = s && (s[0].col !== s[1].col || s[0].row !== s[1].row);
      const d = await documents.pick_save(`${v.wb.name}.csv`);
      if (!d) return;
      const o = await v.cmds.execute("file.exportCsv", many ? { range: v.range() } : {});
      await d.write_text(String(o.output ?? ""));
      print(`sheet: exported ${d.name}`);
      v.status.set(`Exported ${d.name}`);
    }));
  });
}

// Several workbooks (books.ts): each open one a button that makes it the
// one shown; a new one, one opened by the field's name (a saved file, or a
// sample: budget, test, formatted), the active one closed.
export function booksBar(v: SheetView): void {
  const which = state("");
  ui.hstack(() => {
    ui.text("Workbooks:");
    each(v.bookNames, (n: string) => n, (n: string) => ui.button(n, () => {
      ui.onclick(() => v.run("book.activate", { name: n.replace(/ \(active\)$/, "") }));
    }));
    bind(ui.textfield("workbook to open", () => {}), which);
    v.btn("New workbook", () => v.run("book.new"));
    v.btn("Open workbook", () => v.run("book.open", /^(budget|test|formatted)$/i.test(which.value.trim()) ? { sample: which.value.trim() } : { name: which.value }));
    v.btn("Close workbook", () => v.run("book.close"));
  });
}

// Versions kept in a git repository (gitstore.ts): a message, Save version
// (a commit of the workbook's text), the versions newest first, each with
// Diff (against the workbook as it is now, cell by cell, before a change is
// kept) and Restore; Diff last compares the last version and the workbook.
export function versionsBar(v: SheetView): void {
  const message = state(""), versions = state<{ id: string; message: string }[]>([]), changes = state<string[]>([]);
  const listed = async () => {
    const o = await v.run("version.list");
    versions.set(o.ok ? o.output as { id: string; message: string }[] : []);
  };
  const diff = async (a?: string) => {
    const o = await v.run("version.diff", a ? { a } : {});
    const found = o.ok ? o.output as Change[] : [];
    changes.set(o.ok ? [found.length ? `${found.length} changed:` : "No changes", ...found.map(changeLine)] : []);
  };
  ui.wrap(() => {
    ui.on_submit(bind(ui.textfield("version message", () => {}), message), async () => { await v.run("version.save", { message: message.value }); message.set(""); await listed(); });
    v.btn("Save version", async () => { await v.run("version.save", { message: message.value }); message.set(""); await listed(); });
    v.btn("Show versions", listed);
    v.btn("Diff last", () => diff());
  });
  each(versions, (x: { id: string; message: string }) => x.id, (x: { id: string; message: string }) => ui.wrap(() => {
    ui.text(`${x.id.slice(0, 7)} ${x.message.split("\n")[0]}`);
    ui.btn(`Diff ${x.id.slice(0, 7)}`, () => diff(x.id));
    ui.btn(`Restore ${x.id.slice(0, 7)}`, async () => { await v.run("version.restore", { id: x.id }); await listed(); });
  }));
  each(changes, (l: string) => l, (l: string) => ui.text(l));
}
