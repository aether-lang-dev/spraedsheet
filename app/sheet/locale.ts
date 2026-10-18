// Localization (7.9): the interface's words in a table, English and
// Spanish; and Spanish formulas as Excel's Spanish edition takes them in a
// cell and the formula bar: its function names (SUMA, SI, BUSCARV ...), its
// TRUE and FALSE (VERDADERO, FALSO), a decimal comma and ';' between
// arguments. What is stored is always canonical, English with '.' and ',':
// a file, a script, the bus and the clipboard read the same whatever the
// person types in. Pure.
export type Locale = "en" | "es";
export const LOCALES: readonly Locale[] = ["en", "es"];

// Excel's Spanish function names, as its documentation lists them (the
// ones a person types most), by the English.
export const ES_FUNCTIONS: Readonly<Record<string, string>> = {
  SUM: "SUMA", AVERAGE: "PROMEDIO", COUNT: "CONTAR", COUNTA: "CONTARA", COUNTBLANK: "CONTAR.BLANCO", MAX: "MAX", MIN: "MIN",
  IF: "SI", IFERROR: "SI.ERROR", IFS: "SI.CONJUNTO", AND: "Y", OR: "O", NOT: "NO", XOR: "XO",
  SUMIF: "SUMAR.SI", SUMIFS: "SUMAR.SI.CONJUNTO", COUNTIF: "CONTAR.SI", COUNTIFS: "CONTAR.SI.CONJUNTO", AVERAGEIF: "PROMEDIO.SI", AVERAGEIFS: "PROMEDIO.SI.CONJUNTO",
  VLOOKUP: "BUSCARV", HLOOKUP: "BUSCARH", XLOOKUP: "BUSCARX", LOOKUP: "BUSCAR", INDEX: "INDICE", MATCH: "COINCIDIR", XMATCH: "COINCIDIRX", CHOOSE: "ELEGIR", OFFSET: "DESREF", INDIRECT: "INDIRECTO",
  ROUND: "REDONDEAR", ROUNDUP: "REDONDEAR.MAS", ROUNDDOWN: "REDONDEAR.MENOS", INT: "ENTERO", TRUNC: "TRUNCAR", ABS: "ABS", MOD: "RESIDUO", POWER: "POTENCIA", SQRT: "RAIZ", PI: "PI", RAND: "ALEATORIO", RANDBETWEEN: "ALEATORIO.ENTRE", PRODUCT: "PRODUCTO", SUMPRODUCT: "SUMAPRODUCTO",
  TODAY: "HOY", NOW: "AHORA", DATE: "FECHA", YEAR: "AÑO", MONTH: "MES", DAY: "DIA", WEEKDAY: "DIASEM", HOUR: "HORA", MINUTE: "MINUTO", SECOND: "SEGUNDO", EDATE: "FECHA.MES", EOMONTH: "FIN.MES", NETWORKDAYS: "DIAS.LAB", DATEDIF: "SIFECHA",
  CONCATENATE: "CONCATENAR", CONCAT: "CONCAT", TEXTJOIN: "UNIRCADENAS", LEFT: "IZQUIERDA", RIGHT: "DERECHA", MID: "EXTRAE", LEN: "LARGO", UPPER: "MAYUSC", LOWER: "MINUSC", PROPER: "NOMPROPIO", TRIM: "ESPACIOS",
  TEXT: "TEXTO", VALUE: "VALOR", FIND: "ENCONTRAR", SEARCH: "HALLAR", SUBSTITUTE: "SUSTITUIR", REPLACE: "REEMPLAZAR", REPT: "REPETIR", EXACT: "IGUAL",
  ISBLANK: "ESBLANCO", ISNUMBER: "ESNUMERO", ISTEXT: "ESTEXTO", ISERROR: "ESERROR", ISNA: "ESNOD", NA: "NOD",
  MEDIAN: "MEDIANA", MODE: "MODA", STDEV: "DESVEST", VAR: "VAR", LARGE: "K.ESIMO.MAYOR", SMALL: "K.ESIMO.MENOR", RANK: "JERARQUIA",
  PMT: "PAGO", FV: "VF", PV: "VA", NPV: "VNA", IRR: "TIR", RATE: "TASA", NPER: "NPER",
  FILTER: "FILTRAR", SORT: "ORDENAR", UNIQUE: "UNICOS", SEQUENCE: "SECUENCIA", TRANSPOSE: "TRANSPONER", LET: "LET", LAMBDA: "LAMBDA",
};
const EN_FUNCTIONS: Readonly<Record<string, string>> = Object.fromEntries(Object.entries(ES_FUNCTIONS).map(([en, es]) => [es, en]));
const BOOLS: Readonly<Record<string, [string, string]>> = { es: ["VERDADERO", "FALSO"], en: ["TRUE", "FALSE"] };

// A formula's pieces: strings ("..."), sheet names in quotes ('...'),
// numbers, words (names, functions, references, booleans), and the rest.
type Piece = { kind: "string" | "quoted" | "number" | "word" | "other"; text: string };
function pieces(f: string, decimal: "." | ","): Piece[] {
  const out: Piece[] = [];
  let i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      for (; j < f.length; j++) {
        if (f[j] === ch) { if (f[j + 1] === ch) { j++; continue; } break; }
      }
      out.push({ kind: ch === '"' ? "string" : "quoted", text: f.slice(i, j + 1) });
      i = j + 1;
      continue;
    }
    const num = decimal === "," ? /^\d+(?:,\d+)?(?:[eE][-+]?\d+)?%?/.exec(f.slice(i)) : /^\d+(?:\.\d+)?(?:[eE][-+]?\d+)?%?/.exec(f.slice(i));
    const prev = out.length ? out[out.length - 1] : null;
    if (num && !(prev && prev.kind === "word" && /[A-Za-z$]$/.test(prev.text))) {
      out.push({ kind: "number", text: num[0] });
      i += num[0].length;
      continue;
    }
    const word = /^[A-Za-zÀ-ÿ_$][A-Za-zÀ-ÿ0-9_.$]*/.exec(f.slice(i));
    if (word) { out.push({ kind: "word", text: word[0] }); i += word[0].length; continue; }
    out.push({ kind: "other", text: ch });
    i++;
  }
  return out;
}

// What a person typed, in a locale, as it is stored: "=SI(VERDADERO;
// SUMA(1,5;2,25);0)" in Spanish is "=IF(TRUE,SUM(1.5,2.25),0)". Not a
// formula: a number's decimal comma ("1,5" is 1.5), else as typed.
export function canonical(text: string, locale: Locale): string {
  if (locale === "en") return text;
  if (!text.startsWith("=")) return /^\s*-?\d+,\d+\s*$/.test(text) ? text.replace(",", ".") : text;
  return "=" + pieces(text.slice(1), ",").map((p, i, all) => {
    if (p.kind === "number") return p.text.replace(",", ".");
    if (p.kind === "other" && p.text === ";") return ",";
    if (p.kind === "word") {
      const up = p.text.toUpperCase(), call = all[i + 1]?.text === "(";
      if (call && EN_FUNCTIONS[up]) return EN_FUNCTIONS[up];
      if (!call && up === BOOLS.es[0]) return "TRUE";
      if (!call && up === BOOLS.es[1]) return "FALSE";
    }
    return p.text;
  }).join("");
}

// A stored input as a person in a locale reads it: the other way.
export function localized(input: string, locale: Locale): string {
  if (locale === "en") return input;
  if (!input.startsWith("=")) return /^\s*-?\d+\.\d+\s*$/.test(input) ? input.replace(".", ",") : input;
  return "=" + pieces(input.slice(1), ".").map((p, i, all) => {
    if (p.kind === "number") return p.text.replace(".", ",");
    if (p.kind === "other" && p.text === ",") return ";";
    if (p.kind === "word") {
      const up = p.text.toUpperCase(), call = all[i + 1]?.text === "(";
      if (call && ES_FUNCTIONS[up]) return ES_FUNCTIONS[up];
      if (!call && up === "TRUE") return BOOLS.es[0];
      if (!call && up === "FALSE") return BOOLS.es[1];
    }
    return p.text;
  }).join("");
}

// The interface's words: the menus' names and the panels' tabs, the
// status line's, the toolbar's tooltips. An English word not in the table
// is shown as it is.
export const STRINGS: Readonly<Record<Locale, Readonly<Record<string, string>>>> = {
  en: {},
  es: {
    "File": "Archivo", "Edit": "Edición", "View": "Vista", "Insert": "Insertar", "Format": "Formato", "Data": "Datos", "Formulas": "Fórmulas", "Review": "Revisar", "Script": "Script",
    "Looks": "Aspecto", "Names": "Nombres", "Rules": "Reglas", "Validation": "Validación", "Editors": "Editores", "Find & replace": "Buscar y reemplazar", "Audit": "Auditoría", "Charts": "Gráficos", "Lists": "Listas",
    "Panels:": "Paneles:", "Ready": "Listo", "Average": "Promedio", "Count": "Recuento", "Numerical Count": "Recuento numérico", "Min": "Mín.", "Max": "Máx.", "Sum": "Suma",
    "Undo (Ctrl+Z)": "Deshacer (Ctrl+Z)", "Redo (Ctrl+Y)": "Rehacer (Ctrl+Y)", "Cut (Ctrl+X)": "Cortar (Ctrl+X)", "Copy (Ctrl+C)": "Copiar (Ctrl+C)", "Paste (Ctrl+V)": "Pegar (Ctrl+V)",
    "Bold (Ctrl+B)": "Negrita (Ctrl+N)", "Italic (Ctrl+I)": "Cursiva (Ctrl+K)", "Underline (Ctrl+U)": "Subrayado (Ctrl+S)", "Align Left": "Alinear a la izquierda", "Center": "Centrar", "Align Right": "Alinear a la derecha",
    "Comma Style": "Estilo millares", "Currency Style": "Formato de moneda", "Percent Style": "Estilo porcentual", "Borders: an outline round the selection": "Bordes: un contorno alrededor de la selección",
    "Fill Color: yellow": "Color de relleno: amarillo", "Font Color: red": "Color de fuente: rojo", "Sort Ascending, by the selected cell's column": "Orden ascendente, por la columna de la celda",
    "AutoFilter: the selection's rows by the filter field's criterion": "Autofiltro: las filas de la selección por el criterio", "Chart of the selection": "Gráfico de la selección",
  },
};
export const t = (locale: Locale, s: string): string => STRINGS[locale][s] ?? s;
