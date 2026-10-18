// The selected cell's comment and number format (or style: "bold;
// color:red", "$#,##0.00", "percent:0.0%"); Return sets each.
// on_submit gives the field back.
import type { SheetView } from "./view.ts";

export function notesBar(v: SheetView): void {
  ui.wrap(() => {
    const comment = ui.on_submit(bind(ui.textfield("comment", () => {}), v.comment), () => v.run("cell.comment", { cell: v.here(), text: v.comment.value }));
    const format = ui.on_submit(bind(ui.textfield("format", () => {}), v.format), () => v.run("cell.format", { cell: v.here(), format: v.format.value }));
    // The menus' Comment and Number format go to the field.
    v.action("Comment", () => ui.focus(comment));
    v.action("Number format", () => ui.focus(format));
  });
}

// Threaded comments (Review > New Comment): the selected cell's thread
// begun or replied to with the text field's words, resolved or reopened,
// deleted; notes converted; and the comments pane, every thread on the
// sheet, kept under the buttons.
export function commentsBar(v: SheetView): void {
  const words = state(""), pane = state("");
  const listed = async () => {
    const o = await v.run("comment.list");
    pane.set(Array.isArray(o.output) && o.output.length ? (o.output as string[]).join("\n") : "No comments");
  };
  const say = (id: string, params: Record<string, string> = {}) => async () => { await v.run(id, { cell: v.here(), ...params }); await listed(); };
  ui.wrap(() => {
    const field = bind(ui.textfield("comment thread text", () => {}), words);
    v.action("Comment thread field", () => ui.focus(field));
    v.btn("New comment", async () => { await v.run("comment.new", { cell: v.here(), text: words.value }); await listed(); });
    v.btn("Reply", async () => { await v.run("comment.reply", { cell: v.here(), text: words.value }); await listed(); });
    v.btn("Resolve thread", say("comment.resolve"));
    v.btn("Reopen thread", say("comment.resolve", { off: "1" }));
    v.btn("Delete thread", say("comment.delete"));
    v.btn("Convert notes to comments", async () => { await v.run("comment.convertNotes"); await listed(); });
    v.btn("Show comments", listed);
  });
  bind(ui.text(""), pane);
}
