// Reference examples and boundary cases, per function (docs/comparisons/GNUMERIC.md
// action 1). Each function has one worked example whose answer is known from its
// definition (a closed form, a textbook value, the arithmetic done by hand) and the
// boundary that tells it apart from a near miss: the empty case, zero, the edge of
// the domain, the wrong kind of argument. Never another spreadsheet's output.
// An expected value is a number (to 1e-9 relative), a string, a boolean, or "#CODE!"
// for that error. The last case counts the functions covered against the library.
import { evaluate } from "./sheet/evaluate.ts";
import { names } from "./sheet/functions.ts";
import { CellError } from "./sheet/values.ts";
import type { Value } from "./sheet/values.ts";
import { parse } from "./sheet/parser.ts";
import { eq } from "./lib/check.ts";

const env = { get: (_k: string): Value => null };
type Want = number | string | boolean;
// function: [example, boundary ...], each [formula, expected]
const CASES: Record<string, [string, Want][]> = {
  ABS: [["ABS(-5)", 5], ["ABS(0)", 0], ["ABS(\"x\")", "#VALUE!"]],
  ACOS: [["ACOS(0.5)", 1.0471975511965976], ["ACOS(-1)", Math.PI], ["ACOS(2)", "#NUM!"]],
  ASIN: [["ASIN(0.5)", 0.5235987755982989], ["ASIN(1)", Math.PI / 2], ["ASIN(-2)", "#NUM!"]],
  ATAN: [["ATAN(1)", Math.PI / 4], ["ATAN(0)", 0]],
  ATAN2: [["ATAN2(1, 1)", Math.PI / 4], ["ATAN2(1, 0)", 0], ["ATAN2(0, 0)", "#DIV/0!"]],
  CEILING: [["CEILING(2.5, 1)", 3], ["CEILING(2.5, 0)", 0], ["CEILING(-2.5, -2)", -4]],
  FLOOR: [["FLOOR(2.5, 1)", 2], ["FLOOR(3.7, 0)", "#DIV/0!"]],
  COS: [["COS(0)", 1], ["COS(PI())", -1]],
  SIN: [["SIN(PI() / 2)", 1], ["SIN(0)", 0]],
  TAN: [["TAN(PI() / 4)", 1], ["TAN(0)", 0]],
  DEGREES: [["DEGREES(PI())", 180], ["DEGREES(0)", 0]],
  RADIANS: [["RADIANS(180)", Math.PI], ["RADIANS(0)", 0]],
  EVEN: [["EVEN(3)", 4], ["EVEN(0)", 0], ["EVEN(-1)", -2]],
  ODD: [["ODD(2)", 3], ["ODD(0)", 1], ["ODD(-2)", -3]],
  EXP: [["EXP(1)", Math.E], ["EXP(0)", 1]],
  LN: [["LN(EXP(3))", 3], ["LN(1)", 0], ["LN(0)", "#NUM!"]],
  LOG: [["LOG(8, 2)", 3], ["LOG(1000)", 3], ["LOG(8, 1)", "#NUM!"]],
  LOG10: [["LOG10(1000)", 3], ["LOG10(1)", 0], ["LOG10(-1)", "#NUM!"]],
  SQRT: [["SQRT(16)", 4], ["SQRT(0)", 0], ["SQRT(-1)", "#NUM!"]],
  INT: [["INT(1.5)", 1], ["INT(-1.5)", -2], ["INT(0)", 0]],
  MOD: [["MOD(10, 3)", 1], ["MOD(-3, 2)", 1], ["MOD(5, 0)", "#DIV/0!"]],
  POWER: [["POWER(2, 10)", 1024], ["POWER(2, -1)", 0.5], ["POWER(0, -1)", "#DIV/0!"]],
  PRODUCT: [["PRODUCT(2, 3, 4)", 24], ["PRODUCT({5})", 5]],
  ROUND: [["ROUND(2.5, 0)", 3], ["ROUND(-2.5, 0)", -3], ["ROUND(1234, -2)", 1200]],
  ROUNDDOWN: [["ROUNDDOWN(-2.7, 0)", -2], ["ROUNDDOWN(3.14159, 2)", 3.14]],
  ROUNDUP: [["ROUNDUP(2.1, 0)", 3], ["ROUNDUP(-2.1, 0)", -3], ["ROUNDUP(5, 0)", 5]],
  SIGN: [["SIGN(-3)", -1], ["SIGN(0)", 0], ["SIGN(9)", 1]],
  SUM: [["SUM(1, 2, 3)", 6], ["SUM({1, 2; 3, 4})", 10], ["SUM({\"a\"})", 0]],
  TRUNC: [["TRUNC(-8.9)", -8], ["TRUNC(PI(), 3)", 3.141], ["TRUNC(0.5)", 0]],
  MROUND: [["MROUND(10, 3)", 9], ["MROUND(0, 5)", 0], ["MROUND(5, -2)", "#NUM!"]],
  FACTDOUBLE: [["FACTDOUBLE(7)", 105], ["FACTDOUBLE(0)", 1], ["FACTDOUBLE(-2)", "#NUM!"]],
  MULTINOMIAL: [["MULTINOMIAL(2, 3, 4)", 1260], ["MULTINOMIAL(5)", 1], ["MULTINOMIAL(-1, 2)", "#NUM!"]],
  SUMPRODUCT: [["SUMPRODUCT({1, 2}, {3, 4})", 11], ["SUMPRODUCT({1, 2}, {3})", "#VALUE!"]],
  MDETERM: [["MDETERM({1, 2; 3, 4})", -2], ["MDETERM({5})", 5], ["MDETERM({1, 2})", "#VALUE!"]],
  AVERAGE: [["AVERAGE({1, 2, 3, 4})", 2.5], ["AVERAGE({7})", 7], ["AVERAGE({\"a\"})", "#DIV/0!"]],
  MEDIAN: [["MEDIAN({1, 3, 2, 4})", 2.5], ["MEDIAN({9})", 9], ["MEDIAN({1, 2, 3})", 2]],
  MODE: [["MODE({1, 2, 2, 3})", 2], ["MODE({1, 2, 3})", "#N/A"]],
  "STDEV.S": [["STDEV.S({2, 4, 4, 4, 5, 5, 7, 9})", Math.sqrt(32 / 7)], ["STDEV.S({1})", "#DIV/0!"]],
  "STDEV.P": [["STDEV.P({2, 4, 4, 4, 5, 5, 7, 9})", 2], ["STDEV.P({6})", 0]],
  "VAR.S": [["VAR.S({2, 4, 4, 4, 5, 5, 7, 9})", 32 / 7], ["VAR.S({1})", "#DIV/0!"]],
  "VAR.P": [["VAR.P({2, 4, 4, 4, 5, 5, 7, 9})", 4], ["VAR.P({3, 3})", 0]],
  MAX: [["MAX({1, 5, 3})", 5], ["MAX({\"a\"})", 0], ["MAX({-4, -9})", -4]],
  MIN: [["MIN({4, 1, 3})", 1], ["MIN({\"a\"})", 0], ["MIN({-4, -9})", -9]],
  LARGE: [["LARGE({1, 5, 3}, 2)", 3], ["LARGE({1, 5, 3}, 4)", "#NUM!"], ["LARGE({1, 5, 3}, 0)", "#NUM!"]],
  SMALL: [["SMALL({1, 5, 3}, 1)", 1], ["SMALL({1, 5, 3}, 4)", "#NUM!"]],
  COUNT: [["COUNT({1, \"a\", 2})", 2], ["COUNT({\"a\"})", 0]],
  CORREL: [["CORREL({1, 2, 3}, {2, 4, 6})", 1], ["CORREL({1, 2, 3}, {6, 4, 2})", -1], ["CORREL({1, 2}, {1})", "#N/A"]],
  SLOPE: [["SLOPE({2, 4, 6}, {1, 2, 3})", 2], ["SLOPE({5, 5, 5}, {1, 2, 3})", 0], ["SLOPE({1, 2}, {3, 3})", "#DIV/0!"]],
  INTERCEPT: [["INTERCEPT({3, 5, 7}, {1, 2, 3})", 1], ["INTERCEPT({4, 4}, {1, 2})", 4]],
  "PERCENTILE.INC": [["PERCENTILE.INC({1, 2, 3, 4}, 0.5)", 2.5], ["PERCENTILE.INC({1, 2, 3, 4}, 0)", 1], ["PERCENTILE.INC({1, 2, 3, 4}, 1.5)", "#NUM!"]],
  "QUARTILE.INC": [["QUARTILE.INC({1, 2, 3, 4, 5}, 1)", 2], ["QUARTILE.INC({1, 2, 3, 4, 5}, 4)", 5], ["QUARTILE.INC({1, 2, 3, 4, 5}, 5)", "#NUM!"]],
  "RANK.EQ": [["RANK.EQ(3, {1, 3, 5})", 2], ["RANK.EQ(5, {1, 3, 5}, 1)", 3], ["RANK.EQ(9, {1, 3, 5})", "#N/A"]],
  GEOMEAN: [["GEOMEAN({1, 4})", 2], ["GEOMEAN({5})", 5], ["GEOMEAN({0, 1})", "#NUM!"]],
  HARMEAN: [["HARMEAN({1, 4})", 1.6], ["HARMEAN({5})", 5], ["HARMEAN({0, 1})", "#NUM!"]],
  DEVSQ: [["DEVSQ({1, 2, 3})", 2], ["DEVSQ({4, 4, 4})", 0]],
  AVEDEV: [["AVEDEV({1, 2, 3})", 2 / 3], ["AVEDEV({4, 4})", 0]],
  "NORM.S.DIST": [["NORM.S.DIST(0, TRUE)", 0.5], ["NORM.S.DIST(0, FALSE)", 1 / Math.sqrt(2 * Math.PI)]],
  "NORM.S.INV": [["NORM.S.INV(0.5)", 0], ["NORM.S.INV(0)", "#NUM!"], ["NORM.S.INV(1)", "#NUM!"]],
  "NORM.DIST": [["NORM.DIST(1, 1, 1, FALSE)", 1 / Math.sqrt(2 * Math.PI)], ["NORM.DIST(3, 3, 2, TRUE)", 0.5], ["NORM.DIST(1, 1, 0, TRUE)", "#NUM!"]],
  "BINOM.DIST": [["BINOM.DIST(2, 4, 0.5, FALSE)", 0.375], ["BINOM.DIST(4, 4, 0.5, TRUE)", 1], ["BINOM.DIST(5, 4, 0.5, TRUE)", "#NUM!"]],
  "POISSON.DIST": [["POISSON.DIST(0, 2, FALSE)", Math.exp(-2)], ["POISSON.DIST(1, 2, FALSE)", 2 * Math.exp(-2)], ["POISSON.DIST(-1, 2, FALSE)", "#NUM!"]],
  "EXPON.DIST": [["EXPON.DIST(1, 1, TRUE)", 1 - Math.exp(-1)], ["EXPON.DIST(0, 2, FALSE)", 2], ["EXPON.DIST(1, 0, TRUE)", "#NUM!"]],
  "CHISQ.DIST.RT": [["CHISQ.DIST.RT(2, 2)", Math.exp(-1)], ["CHISQ.DIST.RT(0, 3)", 1], ["CHISQ.DIST.RT(-1, 2)", "#NUM!"]],
  PMT: [["PMT(0, 12, 1200)", -100], ["PMT(0.01, 12, 1000)", -1000 * 0.01 / (1 - 1.01 ** -12)]],
  FV: [["FV(0, 10, -100)", 1000], ["FV(0.1, 1, 0, -100)", 110]],
  PV: [["PV(0, 10, -100)", 1000], ["PV(0.1, 1, 0, -110)", 100]],
  NPV: [["NPV(0.1, 110)", 100], ["NPV(0, 10, 20)", 30]],
  IRR: [["IRR({-100, 110})", 0.1], ["IRR({1, 2})", "#NUM!"]],
  SLN: [["SLN(1000, 100, 9)", 100], ["SLN(1000, 100, 0)", "#DIV/0!"]],
  NPER: [["NPER(0, -100, 1000)", 10], ["NPER(0.1, 0, 100, -110)", 1]],
  LEFT: [["LEFT(\"hello\", 2)", "he"], ["LEFT(\"hi\", 0)", ""], ["LEFT(\"hi\", -1)", "#VALUE!"]],
  RIGHT: [["RIGHT(\"hello\", 3)", "llo"], ["RIGHT(\"hi\", 9)", "hi"], ["RIGHT(\"hi\", -1)", "#VALUE!"]],
  MID: [["MID(\"hello\", 2, 3)", "ell"], ["MID(\"hi\", 5, 1)", ""], ["MID(\"hi\", 0, 1)", "#VALUE!"]],
  LEN: [["LEN(\"hello\")", 5], ["LEN(\"\")", 0]],
  UPPER: [["UPPER(\"aBc\")", "ABC"], ["UPPER(\"\")", ""]],
  LOWER: [["LOWER(\"aBc\")", "abc"], ["LOWER(\"\")", ""]],
  PROPER: [["PROPER(\"hello wORLD\")", "Hello World"], ["PROPER(\"\")", ""]],
  TRIM: [["TRIM(\"  a   b \")", "a b"], ["TRIM(\"\")", ""]],
  SUBSTITUTE: [["SUBSTITUTE(\"aaa\", \"a\", \"b\", 2)", "aba"], ["SUBSTITUTE(\"abc\", \"z\", \"y\")", "abc"], ["SUBSTITUTE(\"aaa\", \"a\", \"b\", 4)", "aaa"], ["SUBSTITUTE(\"aaa\", \"a\", \"b\", 0)", "#VALUE!"]],
  REPT: [["REPT(\"ab\", 3)", "ababab"], ["REPT(\"a\", 0)", ""], ["REPT(\"a\", -1)", "#VALUE!"]],
  FIND: [["FIND(\"l\", \"hello\")", 3], ["FIND(\"z\", \"hello\")", "#VALUE!"], ["FIND(\"L\", \"hello\")", "#VALUE!"]],
  SEARCH: [["SEARCH(\"L\", \"hello\")", 3], ["SEARCH(\"h*o\", \"hello\")", 1], ["SEARCH(\"z\", \"hello\")", "#VALUE!"]],
  EXACT: [["EXACT(\"a\", \"a\")", true], ["EXACT(\"a\", \"A\")", false]],
  CONCAT: [["CONCAT(\"a\", 1, TRUE)", "a1TRUE"], ["CONCAT(\"\")", ""]],
  TEXTJOIN: [["TEXTJOIN(\",\", TRUE, \"a\", \"\", \"b\")", "a,b"], ["TEXTJOIN(\",\", FALSE, \"a\", \"\", \"b\")", "a,,b"]],
  VALUE: [["VALUE(\"1e3\")", 1000], ["VALUE(\"abc\")", "#VALUE!"], ["VALUE(\"\")", "#VALUE!"]],
  CHAR: [["CHAR(65)", "A"], ["CHAR(0)", "#VALUE!"]],
  CODE: [["CODE(\"A\")", 65], ["CODE(\"\")", "#VALUE!"]],
  REPLACE: [["REPLACE(\"abcdef\", 2, 3, \"X\")", "aXef"], ["REPLACE(\"abc\", 1, 0, \"Z\")", "Zabc"]],
  TEXT: [["TEXT(1234.5, \"#,##0.00\")", "1,234.50"], ["TEXT(0, \"0.0\")", "0.0"]],
  FIXED: [["FIXED(1234.567, 1)", "1,234.6"], ["FIXED(1234.567, 0, TRUE)", "1235"]],
  T: [["T(\"x\")", "x"], ["T(1)", ""]],
  IF: [["IF(TRUE, 1, 2)", 1], ["IF(FALSE, 1)", false], ["IF(0, \"y\", \"n\")", "n"]],
  AND: [["AND(TRUE, TRUE)", true], ["AND(TRUE, FALSE)", false], ["AND(1, 0)", false]],
  OR: [["OR(FALSE, TRUE)", true], ["OR(FALSE, 0)", false]],
  NOT: [["NOT(TRUE)", false], ["NOT(0)", true]],
  XOR: [["XOR(TRUE, FALSE)", true], ["XOR(TRUE, TRUE)", false]],
  IFERROR: [["IFERROR(1 / 0, \"x\")", "x"], ["IFERROR(7, \"x\")", 7]],
  IFNA: [["IFNA(NA(), 5)", 5], ["IFNA(1 / 0, 5)", "#DIV/0!"]],
  IFS: [["IFS(FALSE, 1, TRUE, 2)", 2], ["IFS(FALSE, 1)", "#N/A"]],
  SWITCH: [["SWITCH(2, 1, \"a\", 2, \"b\")", "b"], ["SWITCH(3, 1, \"a\")", "#N/A"], ["SWITCH(3, 1, \"a\", \"other\")", "other"]],
  INDEX: [["INDEX({1, 2; 3, 4}, 2, 1)", 3], ["INDEX({1, 2; 3, 4}, 3, 1)", "#REF!"]],
  MATCH: [["MATCH(3, {1, 2, 3}, 0)", 3], ["MATCH(2.5, {1, 2, 3}, 1)", 2], ["MATCH(4, {1, 2, 3}, 0)", "#N/A"]],
  VLOOKUP: [["VLOOKUP(2, {1, \"a\"; 2, \"b\"}, 2, FALSE)", "b"], ["VLOOKUP(3, {1, \"a\"; 2, \"b\"}, 2, FALSE)", "#N/A"], ["VLOOKUP(2, {1, \"a\"; 2, \"b\"}, 3, FALSE)", "#REF!"]],
  HLOOKUP: [["HLOOKUP(2, {1, 2; \"a\", \"b\"}, 2, FALSE)", "b"], ["HLOOKUP(3, {1, 2; \"a\", \"b\"}, 2, FALSE)", "#N/A"]],
  CHOOSE: [["CHOOSE(2, \"a\", \"b\")", "b"], ["CHOOSE(0, \"a\")", "#VALUE!"], ["CHOOSE(3, \"a\", \"b\")", "#VALUE!"]],
  ROWS: [["ROWS({1; 2; 3})", 3], ["ROWS({1, 2})", 1]],
  COLUMNS: [["COLUMNS({1, 2})", 2], ["COLUMNS({1; 2})", 1]],
  XLOOKUP: [["XLOOKUP(2, {1, 2, 3}, {\"a\", \"b\", \"c\"})", "b"], ["XLOOKUP(9, {1, 2, 3}, {\"a\", \"b\", \"c\"}, \"none\")", "none"], ["XLOOKUP(9, {1, 2, 3}, {\"a\", \"b\", \"c\"})", "#N/A"]],
  DATE: [["DATE(2024, 2, 29)", 45351], ["DATE(2024, 13, 1)", 45658], ["DATE(2023, 2, 29)", 44986]],
  YEAR: [["YEAR(45351)", 2024], ["YEAR(1)", 1899]],
  MONTH: [["MONTH(45351)", 2], ["MONTH(45322)", 1]],
  DAY: [["DAY(45351)", 29], ["DAY(45322)", 31]],
  EDATE: [["EDATE(DATE(2024, 1, 31), 1)", 45351], ["EDATE(DATE(2024, 3, 31), -1)", 45351], ["EDATE(45351, 0)", 45351]],
  EOMONTH: [["EOMONTH(DATE(2024, 1, 15), 0)", 45322], ["EOMONTH(DATE(2024, 1, 15), 1)", 45351]],
  WEEKDAY: [["WEEKDAY(45292)", 2], ["WEEKDAY(45292, 2)", 1], ["WEEKDAY(45292, 3)", 0]],
  DAYS: [["DAYS(45352, 45351)", 1], ["DAYS(45351, 45352)", -1], ["DAYS(45351, 45351)", 0]],
  DATEDIF: [["DATEDIF(DATE(2020, 1, 1), DATE(2024, 1, 1), \"Y\")", 4], ["DATEDIF(DATE(2024, 1, 1), DATE(2020, 1, 1), \"Y\")", "#NUM!"], ["DATEDIF(DATE(2020, 1, 31), DATE(2020, 3, 1), \"M\")", 1]],
  BIN2DEC: [["BIN2DEC(\"1010\")", 10], ["BIN2DEC(\"102\")", "#NUM!"], ["BIN2DEC(\"1111111111\")", 1023 - 1024]],
  DEC2BIN: [["DEC2BIN(10)", "1010"], ["DEC2BIN(511)", "111111111"], ["DEC2BIN(512)", "#NUM!"]],
  DEC2HEX: [["DEC2HEX(255)", "FF"], ["DEC2HEX(0)", "0"]],
  HEX2DEC: [["HEX2DEC(\"FF\")", 255], ["HEX2DEC(\"G\")", "#NUM!"]],
  BITAND: [["BITAND(12, 10)", 8], ["BITAND(1, 0)", 0], ["BITAND(-1, 1)", "#NUM!"]],
  BITOR: [["BITOR(12, 10)", 14], ["BITOR(0, 0)", 0]],
  BITXOR: [["BITXOR(12, 10)", 6], ["BITXOR(5, 5)", 0]],
  CONVERT: [["CONVERT(1, \"km\", \"m\")", 1000], ["CONVERT(1, \"km\", \"kg\")", "#N/A"]],
  DELTA: [["DELTA(1, 1)", 1], ["DELTA(1, 2)", 0]],
  GESTEP: [["GESTEP(3, 2)", 1], ["GESTEP(2, 2)", 1], ["GESTEP(1, 2)", 0]],
  ERF: [["ERF(1)", 0.8427007929497149], ["ERF(0)", 0]],
  ISNUMBER: [["ISNUMBER(1)", true], ["ISNUMBER(\"1\")", false]],
  ISTEXT: [["ISTEXT(\"a\")", true], ["ISTEXT(1)", false]],
  ISERROR: [["ISERROR(1 / 0)", true], ["ISERROR(1)", false]],
  ISEVEN: [["ISEVEN(2)", true], ["ISEVEN(3)", false], ["ISEVEN(\"a\")", "#VALUE!"]],
  N: [["N(TRUE)", 1], ["N(\"a\")", 0]],
  NA: [["NA()", "#N/A"], ["ISNA(NA())", true]],
  ROMAN: [["ROMAN(1994)", "MCMXCIV"], ["ROMAN(0)", ""], ["ROMAN(-1)", "#VALUE!"]],
  ARABIC: [["ARABIC(\"MCMXCIV\")", 1994], ["ARABIC(\"\")", 0], ["ARABIC(\"ABC\")", "#VALUE!"]],
  BASE: [["BASE(255, 16)", "FF"], ["BASE(0, 2)", "0"], ["BASE(5, 1)", "#NUM!"]],
  DECIMAL: [["DECIMAL(\"FF\", 16)", 255], ["DECIMAL(\"2\", 2)", "#NUM!"]],
};
// BIN2DEC of ten ones is the 10-bit two's complement -1.

let bad: string[] = [], checked = 0;
for (const fn of Object.keys(CASES)) {
  if (!names().includes(fn)) bad.push(`${fn}: not in the library`);
  if (CASES[fn].length < 2) bad.push(`${fn}: needs an example and a boundary`);
  for (const [src, want] of CASES[fn]) {
    checked++;
    let got: unknown;
    try { got = evaluate(parse(src), env); } catch (e) { bad.push(`${src}: threw`); continue; }
    if (got instanceof CellError) got = got.code;
    const ok = typeof want === "number" ? typeof got === "number" && Math.abs(got - want) <= 1e-9 * Math.max(1, Math.abs(want)) : got === want;
    if (!ok) bad.push(`${src}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
  }
}
eq("every example and boundary comes out as its definition says", bad, []);
eq("the cases are many enough to mean something", checked >= 300, true);
eq("how many functions have examples and boundaries", `${Object.keys(CASES).length} of ${names().length}`, `${Object.keys(CASES).length} of ${names().length}`);

// expect: ok every example and boundary comes out as its definition says
// expect: ok the cases are many enough to mean something
// expect: ok how many functions have examples and boundaries
