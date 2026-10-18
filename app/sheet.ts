"seeks documents, reduced functionality without";
"seeks clipboard, reduced functionality without";
"seeks control-bus, reduced functionality without";
"seeks open-urls, reduced functionality without";
"seeks print, reduced functionality without";
// Spreadsheet (sae/docs/roadmap.md demo 18), declared as panels over a seam.
//
// app/sheet/ is the spreadsheet with no ui: two formula languages (Excel's,
// and RexxJS's as the Tauri spreadsheet has them), a Sheet that keeps
// who-reads-whom and works out what changed in order, notes and formats,
// and under them all the commands (commands.ts): every action is one, with
// string parameters and a journal, which RexxJS scripts drive through
// ADDRESS SHEET and specs drive headless. app/sheet/view/ is the ui: each
// panel reads the workbook through state and runs commands, and nothing
// else. This page only says which panels, in which order. Run as an
// installed app (sae --app site, app/app.json) it may also ask the person
// for files (documents); in the browser it does without.
import { SheetView } from "./sheet/view/view.ts";
import { languageBar } from "./sheet/view/languagebar.ts";
import { formulaBar } from "./sheet/view/formulabar.ts";
import { commentsBar, notesBar } from "./sheet/view/notesbar.ts";
import { lookBar } from "./sheet/view/lookbar.ts";
import { viewBar } from "./sheet/view/viewbar.ts";
import { namesBar } from "./sheet/view/namesbar.ts";
import { validationBar } from "./sheet/view/validationbar.ts";
import { editorBar } from "./sheet/view/editorbar.ts";
import { sheetGrid, splitPane } from "./sheet/view/grid.ts";
import { sheetButtons, sheetTabs } from "./sheet/view/sheettabs.ts";
import { editBar, protectBar } from "./sheet/view/editbar.ts";
import { findBar } from "./sheet/view/findbar.ts";
import { auditBar } from "./sheet/view/auditbar.ts";
import { rulesBar } from "./sheet/view/rulesbar.ts";
import { chartPanel } from "./sheet/view/chartpanel.ts";
import { dataBar } from "./sheet/view/databar.ts";
import { insertBar } from "./sheet/view/insertbar.ts";
import { commandBar, macroBar } from "./sheet/view/commandbar.ts";
import { booksBar, fileBar } from "./sheet/view/filebar.ts";
import { functionList, problemList, setupScript, statusBar, statusLine } from "./sheet/view/lists.ts";
import { menuBar } from "./sheet/view/menubar.ts";
import { toolBar } from "./sheet/view/toolbar.ts";
import { panelTabs } from "./sheet/view/paneltabs.ts";

// What this page needs of the sae it runs on (its ui API): an older sae
// build than the spreadsheet (sae's ./build.sh not run since a pull) lacks
// some, and the page then stopped partway, its grid half wired (a click on
// it scrolled the page away and typing did nothing). Each one missing is
// said, on the page and the console, and stood in for where it can be.
const NEEDS = ["wrap", "datagrid_widget", "datagrid_show_header", "datagrid_set_row_heights", "datagrid_cell_picture", "datagrid_selection_glide", "context_menu", "on_double_click"];
const lacking = NEEDS.filter((n) => typeof (ui as any)[n] !== "function");
if (typeof (ui as any).wrap !== "function") (ui as any).wrap = ui.hstack;
if (typeof (ui as any).datagrid_widget !== "function") (ui as any).datagrid_widget = () => 0;
// No right-click menus or double clicks on a tab: the sheet buttons still do it all.
if (typeof (ui as any).context_menu !== "function") (ui as any).context_menu = () => {};
if (typeof (ui as any).on_double_click !== "function") (ui as any).on_double_click = () => {};
if (lacking.length) print(`sheet: this sae lacks ui.${lacking.join(", ui.")}: rebuild it (../sae, its ./build.sh)`);

const sheet = new SheetView();

// The page in one column with the window's own padding round it.
ui.vstack(() => {
ui.margin(8, 12, 8, 12);
ui.text("Spreadsheet");
if (lacking.length) ui.text(`This sae is older than the spreadsheet: it lacks ui.${lacking.join(", ui.")}. Rebuild it (../sae, its ./build.sh).`);
toolBar(sheet);
ui.divider();
// The Name Box and the formula bar, one row under the toolbar.
formulaBar(sheet);
statusLine(sheet);
ui.divider();
sheetGrid(sheet);
splitPane(sheet);
sheetTabs(sheet);
statusBar(sheet);
ui.divider();
// Behind the Extra controls (collapsed as it starts, saved with the
// workbook): the sheet buttons and the panels' tabs. A panel opened (its
// tab, the View menu, an action of its) shows either way.
sheet.extras(() => { sheetButtons(sheet); panelTabs(sheet); });
// The bars, as panels under the grid, one open at a time (view.panel).
sheet.panel("edit", () => { editBar(sheet); protectBar(sheet); });
sheet.panel("looks", () => { lookBar(sheet); notesBar(sheet); commentsBar(sheet); });
sheet.panel("view", () => viewBar(sheet));
sheet.panel("names", () => namesBar(sheet));
sheet.panel("rules", () => rulesBar(sheet));
sheet.panel("validation", () => validationBar(sheet));
sheet.panel("editors", () => editorBar(sheet));
sheet.panel("data", () => dataBar(sheet));
sheet.panel("find", () => findBar(sheet));
sheet.panel("audit", () => auditBar(sheet));
sheet.panel("insert", () => insertBar(sheet));
sheet.panel("chart", () => chartPanel(sheet));
sheet.panel("file", () => { fileBar(sheet); booksBar(sheet); });
sheet.panel("script", () => { languageBar(sheet); commandBar(sheet); macroBar(sheet); setupScript(sheet); });
sheet.panel("lists", () => { functionList(sheet); problemList(sheet); });
});
// The menus run the panels' actions: made once the panels are.
menuBar(sheet);
sheet.refresh();

// The control bus (an installed app started with --control-bus): another
// program drives the sheet as the Tauri app's clients do, each request a
// string of REXX (SETCELL("A1", "100")), answered with what it came to.
// A request sae stopped (past the page's time limit, or unanswered at the
// bus's deadline) is told to the commands, which put back what it did.
if (typeof bus !== "undefined") {
  sheet.bus = true;
  sheet.follow();
  bus.serve(async (code: string) => {
    const o = await sheet.cmds.spreadsheet(code, (line) => print(`sheet: ${line}`));
    sheet.refresh();
    if (!o.ok) throw new Error(o.message);
    return o.output ?? null;
  }, async (reason: string) => {
    const what = await sheet.cmds.stopped(reason);
    sheet.refresh();
    print(`sheet: a control-bus request ${reason === "interrupted" ? "was stopped" : "ran past the deadline"}: ${what}`);
  });
}
