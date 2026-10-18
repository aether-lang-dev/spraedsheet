// The workbook's sheets as tabs under the grid, as Excel's: one tab each,
// the active one marked; a click shows a sheet, a tab dragged among the
// others moves its sheet there (sae's ui.reorderable over the row each()
// makes), a double click renames it in place (Return keeps the name,
// Escape or an empty name keeps the old one), and its right-click menu
// (ui.context_menu) has what the Sheet menu has: Rename, Move left and
// right, Copy, Hide, Unhide, a tab colour, Delete. The "+" tab adds a
// sheet. At the row's end the Extra controls toggle (view.extraControls,
// saved with the workbook) shows the rows below: the sheet buttons (a
// name to add a sheet by, rename the active one to or colour its tab
// with) and the panels' tabs.
import type { SheetView } from "./view.ts";

const TAB_COLOURS = ["red", "orange", "green", "blue", "purple"];

export function sheetTabs(v: SheetView): void {
  // The tab being renamed in place, or "".
  const renaming = state("");
  type Tab = { name: string; color: string | null; active: boolean; editing: boolean };
  const shown = computed(() => v.tabs.value.map((t): Tab => ({ ...t, active: t.name === v.active.value, editing: t.name === renaming.value })), v.tabs, v.active, renaming);
  const rgb = (c: string) => [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16) / 255);
  const rename = (from: string, to: string) => {
    renaming.set("");
    const name = to.trim();
    if (name && name !== from) v.run("sheet.rename", { name: from, to: name });
  };
  const menuOf = (name: string) => [
    { label: "Rename", action: () => renaming.set(name) },
    { label: "Move left", action: () => v.run("sheet.move", { name, way: "left" }) },
    { label: "Move right", action: () => v.run("sheet.move", { name, way: "right" }) },
    { label: "Copy", action: () => v.run("sheet.copy", { name }) },
    { label: "Hide", action: () => v.run("sheet.hide", { name }) },
    { label: "Unhide", enabled: v.wb.pages().some((x) => x.hiddenSheet), action: async () => { for (const p of v.wb.pages().filter((x) => x.hiddenSheet)) await v.run("sheet.hide", { name: p.name, off: "1" }); } },
    "-",
    ...TAB_COLOURS.map((c) => ({ label: `Tab color ${c}`, action: () => v.run("sheet.color", { name, color: c }) })),
    { label: "No tab color", action: () => v.run("sheet.color", { name, color: "none" }) },
    "-",
    { label: "Delete", enabled: v.tabs.value.length > 1, action: () => v.run("sheet.delete", { name }) },
  ];
  ui.hstack(2, () => {
    ui.reorderable(
      each(shown, (t: Tab) => `${t.name}|${t.color ?? ""}|${t.active ? 1 : 0}|${t.editing ? 1 : 0}`, (t: Tab) => {
        if (t.editing) {
          const f = ui.textfield("sheet name", () => {});
          ui.set_text(f, t.name);
          ui.a11y_label(f, `Rename ${t.name}`);
          ui.on_submit(f, (s: string) => rename(t.name, s));
          ui.focus(f);
          return f;
        }
        const tab = ui.button(t.name, () => {
          if (t.color) { const [r, g, b] = rgb(t.color); ui.bg_color(r, g, b, 1); }
          else if (t.active) ui.bg_color(1, 1, 1, 1);
          ui.onclick(() => v.run("sheet.activate", { name: t.name }));
        });
        ui.a11y_label(tab, t.active ? `${t.name}, the sheet shown` : t.name);
        if (t.active) ui.add_class(tab, "sheet-tab-active");
        ui.on_double_click(tab, () => renaming.set(t.name));
        ui.context_menu(tab, menuOf(t.name));
        return tab;
      }),
      (from: number, to: number) => v.run("sheet.move", { name: v.tabs.value[from].name, to: String(v.sheets.value.indexOf(v.tabs.value[to].name) + 1) }),
    );
    const plus = ui.btn("+", () => v.run("sheet.add", {}));
    ui.a11y_label(plus, "Add a sheet");
    ui.tooltip(plus, "Add a sheet");
    ui.spacer();
    v.extraToggle = ui.toggle("Extra controls", (on: number) => {
      if ((on ? true : false) !== v.wb.extraControls) v.run("view.extraControls", { on: on ? "1" : "0" });
    });
    ui.tooltip(v.extraToggle, "The rows under the tabs: the sheet buttons and the panels");
  });
}

// The sheet buttons, behind the Extra controls: a name to add a sheet by,
// rename the active one to, or colour its tab with (a colour's name or
// #rrggbb), and the Sheet menu's actions on the active sheet. A formula
// reads another sheet's cells as Costs!B2 (Excel) or Costs_B2 (RexxJS).
export function sheetButtons(v: SheetView): void {
  const name = state("");
  ui.hstack(() => {
    bind(ui.text(""), computed(() => `On ${v.active.value}:`, v.active));
    bind(ui.textfield("sheet", () => {}), name);
    v.btn("Add sheet", () => v.run("sheet.add", name.value.trim() ? { name: name.value } : {}));
    v.btn("Rename sheet", () => v.run("sheet.rename", { name: v.active.value, to: name.value }));
    v.btn("Move left", () => v.run("sheet.move", { name: v.active.value, way: "left" }));
    v.btn("Move right", () => v.run("sheet.move", { name: v.active.value, way: "right" }));
    v.btn("Copy sheet", () => v.run("sheet.copy", name.value.trim() ? { name: v.active.value, to: name.value } : { name: v.active.value }));
    v.btn("Hide sheet", () => v.run("sheet.hide", { name: v.active.value }));
    v.btn("Unhide sheets", async () => { for (const p of v.wb.pages().filter((x) => x.hiddenSheet)) await v.run("sheet.hide", { name: p.name, off: "1" }); });
    v.btn("Tab color", () => v.run("sheet.color", { name: v.active.value, color: name.value.trim() || "none" }));
    v.btn("Delete sheet", () => v.run("sheet.delete", { name: v.active.value }));
  });
}
