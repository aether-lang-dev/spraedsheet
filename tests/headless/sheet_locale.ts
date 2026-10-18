// Localization (7.9): Spanish formulas as typed (Excel's Spanish names,
// VERDADERO and FALSO, a decimal comma, ';' between arguments) stored
// canonical and shown back in Spanish; strings and sheet names left alone;
// a number with a decimal comma; English untouched; the interface's words;
// SETLOCALE, GETLOCALE and LOCALIZEFORMULA from REXX.
import { canonical, localized, t } from "./sheet/locale.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("the reference's example, stored canonical", canonical("=SI(VERDADERO;SUMA(1,5;2,25);0)", "es"), "=IF(TRUE,SUM(1.5,2.25),0)");
eq("and shown back", localized("=IF(TRUE,SUM(1.5,2.25),0)", "es"), "=SI(VERDADERO;SUMA(1,5;2,25);0)");
eq("strings, quoted sheet names and references left alone", [canonical('=SI(A1>2;"uno; dos";\'Hoja 1\'!B2)', "es"), canonical("=BUSCARV(A1;Tabla;2;FALSO)", "es"), canonical("=SUMA(A1:B10)*2,5", "es")],
  ['=IF(A1>2,"uno; dos",\'Hoja 1\'!B2)', "=VLOOKUP(A1,Tabla,2,FALSE)", "=SUM(A1:B10)*2.5"]);
eq("a function name in any case, one with no Spanish name as it is", [canonical("=suma(1;2)", "es"), canonical("=XNPV(0,1;B1:B3;C1:C3)", "es")], ["=SUM(1,2)", "=XNPV(0.1,B1:B3,C1:C3)"]);
eq("a number typed with a decimal comma", [canonical("3,25", "es"), canonical("1.234,5", "es"), canonical("texto, con coma", "es"), localized("3.25", "es")], ["3.25", "1.234,5", "texto, con coma", "3,25"]);
eq("English untouched", [canonical("=SUM(1.5,2)", "en"), localized("=SUM(1.5,2)", "en")], ["=SUM(1.5,2)", "=SUM(1.5,2)"]);
eq("the interface's words", [t("es", "Ready"), t("es", "Looks"), t("es", "Something else"), t("en", "Ready")], ["Listo", "Aspecto", "Something else", "Ready"]);

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await cmds.run("view.locale lang=es");
await cmds.run(`cell.set cell=A1 input="${wb.enteredInput("A1", "=SUMA(1,5;2,25)").replace(/"/g, '""')}"`);
eq("typed in Spanish, worked out, stored canonical, shown in Spanish", [show(wb.sheet.value("A1")), wb.sheet.input("A1"), wb.shownInput("A1")], ["3.75", "=SUM(1.5,2.25)", "=SUMA(1,5;2,25)"]);
eq("refused", (await cmds.run("view.locale lang=fr")).message, "fr is not a locale here (en, es)");
const said: string[] = [];
await cmds.script('SAY GETLOCALE()\nSAY LOCALIZEFORMULA("=PROMEDIO(1;2,5)")\nSAY LOCALIZEFORMULA("=AVERAGE(1,2.5)", "es", "local")\nSAY SETLOCALE("en") GETLOCALE()', (l) => said.push(l));
eq("from REXX", said, ["es", "=AVERAGE(1,2.5)", "=PROMEDIO(1;2,5)", "en en"]);

// expect: ok the reference's example, stored canonical
// expect: ok and shown back
// expect: ok strings, quoted sheet names and references left alone
// expect: ok a function name in any case, one with no Spanish name as it is
// expect: ok a number typed with a decimal comma
// expect: ok English untouched
// expect: ok the interface's words
// expect: ok typed in Spanish, worked out, stored canonical, shown in Spanish
// expect: ok refused
// expect: ok from REXX
