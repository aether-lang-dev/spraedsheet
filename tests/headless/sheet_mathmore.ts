// Maths and logical values found missing or different by IronCalc's calc_tests
// workbooks (docs/comparisons/IRONCALC.md action 2): TRUE and FALSE typed into a
// cell are logical, ROMAN's forms, the reciprocal trigonometric functions, matrices,
// powers and overflow as Excel gives them.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
for (const l of ["cell.set cell=A1 input=TRUE", "cell.set cell=A2 input=false", "cell.set cell=A3 input=Truely", "cell.set cell=B1 input=2", "cell.set cell=B2 input=4", "cell.set cell=C1 input=1", "cell.set cell=C2 input=3"]) await cmds.run(l);
const f = async (formula: string) => { await cmds.run(`cell.set cell=Z1 input=${formula}`); return show(wb.sheet.value("Z1")); };
const row = async (...fs: string[]) => { const out: string[] = []; for (const x of fs) out.push(await f(x)); return out; };

eq("TRUE and FALSE typed are logical (in any case), other text is not", [typeof wb.sheet.value("A1"), typeof wb.sheet.value("A2"), typeof wb.sheet.value("A3")], ["boolean", "boolean", "string"]);
eq("a logical cell is 1 or 0 to the maths functions, not to the engineering ones", await row("=A1+1", "=SQRT(A1)", "=COMBIN(B2,A1)", "=ERF(A1)", "=GESTEP(A1)", "=ISODD(A1)"), ["2", "1", "4", "#VALUE!", "#VALUE!", "#VALUE!"]);

eq("ROMAN, its forms", await row("=ROMAN(499)", "=ROMAN(499,1)", "=ROMAN(499,2)", "=ROMAN(499,3)", "=ROMAN(499,4)", "=ROMAN(469,2)", "=ROMAN(2596,2)", "=ROMAN(0)", "=ROMAN(3999)", "=ROMAN(4000)", "=ROMAN(12,TRUE)", "=ROMAN(12,FALSE)"),
  ["CDXCIX", "LDVLIV", "XDIX", "VDIV", "ID", "LDXIX", "MMDVCI", "", "MMMCMXCIX", "#VALUE!", "XII", "XII"]);
eq("ARABIC", await row('=ARABIC("MCMXCIV")', '=ARABIC("-iv")', '=ARABIC("")', '=ARABIC("ABC")'), ["1994", "-4", "0", "#VALUE!"]);
eq("the reciprocal trigonometric functions", await row("=ROUND(SEC(1),6)", "=ROUND(CSC(1),6)", "=ROUND(COT(1),6)", "=COT(0)", "=ROUND(ACOT(1),6)", "=ROUND(SECH(1),6)", "=ACOTH(1)", "=ROUND(ACOTH(2),6)", "=CSCH(0)"),
  ["1.850816", "1.188395", "0.642093", "#DIV/0!", "0.785398", "0.648054", "#NUM!", "0.549306", "#DIV/0!"]);
eq("counting", await row("=COMBINA(3,2)", "=COMBINA(0,0)", "=PERMUTATIONA(3,2)", "=FACTDOUBLE(7)", "=FACTDOUBLE(-1)", "=FACTDOUBLE(-2)", "=MULTINOMIAL(2,3,4)", "=ROUND(SQRTPI(2),6)", "=SERIESSUM(2,0,1,{1,2,3})", "=SERIESSUM(0,0,1,{1,2})"),
  ["6", "1", "9", "105", "1", "#NUM!", "1260", "2.506628", "17", "#NUM!"]);
eq("sums over pairs", await row("=SUMX2MY2(B1:B2,C1:C2)", "=SUMX2PY2(B1:B2,C1:C2)", "=SUMXMY2(B1:B2,C1:C2)", "=SUMXMY2(B1:B2,C1:C1)"), ["10", "30", "2", "#N/A"]);
eq("the A forms of the spread", await row("=ROUND(STDEVA(B1:B2),6)", "=ROUND(VARPA(B1:B2),6)", "=ROUND(SKEW.P(1,2,6),6)"), ["1.414214", "1", "0.59517"]);
await cmds.run("cell.set cell=E1 input==MMULT(B1:B2,TRANSPOSE(C1:C2))");
await cmds.run("cell.set cell=I1 input==MINVERSE(E1:F2)");
await cmds.run("cell.set cell=K1 input==MUNIT(2)");
eq("matrices", [show(wb.sheet.value("E1")), show(wb.sheet.value("F2")), await f("=MDETERM(E1:F2)"), await f("=MDETERM(MUNIT(3))"), show(wb.sheet.value("K1")), show(wb.sheet.value("L2")), await f("=MMULT(B1:B2,B1:B2)")],
  ["2", "12", "0", "1", "1", "1", "#VALUE!"]);
eq("a singular matrix has no inverse", show(wb.sheet.value("I1")), "#NUM!");

eq("powers and overflow as Excel gives them", await row("=0^0", "=POWER(0,-1)", "=(-8)^(1/3)", "=2^10", "=1E308*10", "=1E308+1E308", "=SUM(1E308,1E308)"), ["#NUM!", "#DIV/0!", "#NUM!", "1024", "#NUM!", "#NUM!", "#NUM!"]);
eq("a criterion that is a logical value matches logical cells", await row('=COUNTIF(A1:A3,TRUE)', '=COUNTIF(A1:A3,FALSE)', '=COUNTIF(A1:A3,"Truely")'), ["1", "1", "1"]);
eq("the binomial and Poisson distributions for a huge count are quick", await row("=BINOM.DIST(500000,1000000,0.5,TRUE)", "=BINOM.INV(1000000,0.5,0.5)", "=ROUND(POISSON.DIST(100000,100000,TRUE),4)", "=BINOM.DIST.RANGE(1000000,0.5,499990,500010)"), ["0.500398943", "500000", "0.5008", "0.01675434433"]);

// expect: ok TRUE and FALSE typed are logical (in any case), other text is not
// expect: ok a logical cell is 1 or 0 to the maths functions, not to the engineering ones
// expect: ok ROMAN, its forms
// expect: ok ARABIC
// expect: ok the reciprocal trigonometric functions
// expect: ok counting
// expect: ok sums over pairs
// expect: ok the A forms of the spread
// expect: ok matrices
// expect: ok a singular matrix has no inverse
// expect: ok powers and overflow as Excel gives them
// expect: ok a criterion that is a logical value matches logical cells
// expect: ok the binomial and Poisson distributions for a huge count are quick
