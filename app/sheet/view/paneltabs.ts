// The panels under the grid as a row of tabs, one open at a time (as
// Excel's task panes, or old Excel's toolbars shown on demand): a tab opens
// its panel, closing the one open; the open one's tab closes it.
// Their words in the locale's language (locale.ts).
import { PANELS } from "../menus.ts";
import { t } from "../locale.ts";
import type { SheetView } from "./view.ts";

export function panelTabs(v: SheetView): void {
  ui.hstack(() => {
    bind(ui.text(""), computed(() => t(v.locale.value, "Panels:"), v.locale));
    for (const p of PANELS) {
      const tab = ui.tooltip(ui.btn(p.label, () => v.run("view.panel", { name: v.wb.panel === p.name ? "none" : p.name })), `The ${p.label} panel: open, or closed when open`);
      bind(tab, computed(() => t(v.locale.value, p.label), v.locale));
    }
  });
}
