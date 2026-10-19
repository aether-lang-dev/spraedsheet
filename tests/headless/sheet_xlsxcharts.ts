// Charts in XLSX (5.9): every kind the app has written as a DrawingML chart
// part (a classic one, or Office 2016's extended one for the histogram,
// waterfall, funnel, treemap, sunburst and box), their series as references to
// the sheet's cells with the values cached; read back, our own charts as they
// were (the kind, the range, the title, where, the options, and drawn the
// same); Excel's own chart parts read in where the kind is modelled, and the
// ones that are not said in the report.
import { unzip, utf8Bytes, utf8Text, zip } from "./sheet/zip.ts";
import { fromBase64, toBase64 } from "./sheet/csv.ts";
import { parseXml, all, kid, attr } from "./sheet/xml.ts";
import { CHART_KINDS } from "./sheet/charts.ts";
import type { ChartKind } from "./sheet/charts.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { readXlsx } from "./sheet/xlsx.ts";
import { parseRef } from "./sheet/xlsxcharts.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const set = async (cells: [string, string][]) => { for (const [k, v] of cells) await cmds.run(`cell.set cell=${k} input="${v}"`); };
// Four quarters of three regions' sales, with a Total to close; the numbers
// of an x and a y and a size, the high, low and close of a stock; a two-level list.
await cmds.run("formulas.use language=Excel");
await set([["B1", "North"], ["C1", "South"], ["D1", "East"],
  ["A2", "Q1"], ["A3", "Q2"], ["A4", "Q3"], ["A5", "Q4"], ["A6", "Total"],
  ["B2", "10"], ["B3", "20"], ["B4", "15"], ["B5", "30"], ["B6", "75"],
  ["C2", "5"], ["C3", "12"], ["C4", "18"], ["C5", "9"], ["C6", "44"],
  ["D2", "8"], ["D3", "7"], ["D4", "22"], ["D5", "14"], ["D6", "51"],
  ["F1", "x"], ["G1", "y1"], ["H1", "y2"], ["F2", "1"], ["F3", "2"], ["F4", "3"], ["F5", "4"], ["F6", "5"],
  ["G2", "2"], ["G3", "4"], ["G4", "5"], ["G5", "4"], ["G6", "6"], ["H2", "1"], ["H3", "3"], ["H4", "2"], ["H5", "5"], ["H6", "4"],
  ["J1", "x"], ["K1", "y"], ["L1", "size"], ["J2", "1"], ["J3", "2"], ["J4", "3"], ["J5", "4"], ["J6", "5"],
  ["K2", "3"], ["K3", "5"], ["K4", "2"], ["K5", "6"], ["K6", "4"], ["L2", "10"], ["L3", "30"], ["L4", "20"], ["L5", "50"], ["L6", "40"],
  ["O2", "Mon"], ["O3", "Tue"], ["O4", "Wed"], ["O5", "Thu"], ["O6", "Fri"], ["P1", "Open"], ["Q1", "High"], ["R1", "Low"], ["S1", "Close"],
  ["P2", "10"], ["P3", "12"], ["P4", "11"], ["P5", "13"], ["P6", "12"], ["Q2", "13"], ["Q3", "14"], ["Q4", "13"], ["Q5", "15"], ["Q6", "14"],
  ["R2", "9"], ["R3", "11"], ["R4", "10"], ["R5", "12"], ["R6", "10"], ["S2", "12"], ["S3", "11"], ["S4", "13"], ["S5", "12"], ["S6", "13"],
  ["X1", "Group"], ["Y1", "Item"], ["Z1", "Size"], ["X2", "Fruit"], ["Y2", "Apple"], ["Z2", "8"], ["Y3", "Pear"], ["Z3", "5"], ["X4", "Veg"], ["Y4", "Leek"], ["Z4", "6"], ["Y5", "Kale"], ["Z5", "3"],
  ["AB1", "A"], ["AC1", "B"], ["AB2", "1"], ["AB3", "2"], ["AB4", "2"], ["AB5", "3"], ["AB6", "9"], ["AC2", "4"], ["AC3", "5"], ["AC4", "6"], ["AC5", "6"], ["AC6", "7"]]);

// Each kind over the data that suits it, and how it is dressed.
const OVER: Record<string, string> = { scatter: "F1:H6", scatterLines: "F1:H6", bubble: "J1:L6", stock: "O1:S6", treemap: "X1:Z5", sunburst: "X1:Z5", box: "AB1:AC6",
  pie: "A1:B6", doughnut: "A1:B6", histogram: "A1:B6", waterfall: "A1:B6", funnel: "A1:B6" };
const message = async (line: string) => (await cmds.run(line)).message;
const kinds = CHART_KINDS as readonly ChartKind[];
for (const [i, kind] of kinds.entries()) {
  await cmds.run(`chart.insert kind=${kind} range=${OVER[kind] ?? "A1:D6"} title="${kind} of sales" at=AF${2 + i * 14}`);
  // Every other chart dressed: axis titles, the legend top, labels, no gridlines.
  if (i % 2 === 0) await cmds.run(`chart.option index=${i + 1} xTitle=Quarter yTitle=Units legend=top labels=1 gridlines=0`);
}
await cmds.run("chart.option index=4 legend=none");
await cmds.run(`chart.option index=${kinds.indexOf("combo") + 1} secondary=3`);
await cmds.run(`chart.option index=${kinds.indexOf("line") + 1} byRows=1`);
await cmds.run(`chart.trendline index=${kinds.indexOf("line") + 1} series=1 type=polynomial order=3 equation=1 r2=1`);
await cmds.run(`chart.trendline index=${kinds.indexOf("scatter") + 1} series=2 type=movingAverage period=3`);
await cmds.run(`chart.trendline index=${kinds.indexOf("bar") + 1} series=1 type=linear equation=1`);
const sheet = wb.page(wb.active)!;
eq("the charts made, one of each kind", sheet.charts.map((c) => c.kind), [...kinds]);

// What the cells hold, the kinds of the drawing, drawn.
const drawn = (w: Workbook) => w.charts().map((c) => JSON.stringify([c.kind, c.title, c.where, c.range, c.shapes]));
const before = drawn(wb);

const o = await cmds.run("file.exportXlsx");
const archive = unzip(fromBase64(String(o.output)));
const names = [...archive.keys()];
const classic = names.filter((n) => /^xl\/charts\/chart\d+\.xml$/.test(n)), extended = names.filter((n) => /^xl\/charts\/chartEx\d+\.xml$/.test(n));
eq("a part for each: eighteen classic, six extended", [classic.length, extended.length], [18, 6]);
eq("said in the report: not that the charts are left out", String(o.message).includes("not written"), false);
const types = utf8Text(archive.get("[Content_Types].xml")!);
eq("the parts' types", [types.includes('PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"'),
  types.includes('PartName="/xl/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"')], [true, true]);
const rels = utf8Text(archive.get("xl/drawings/_rels/drawing1.xml.rels")!);
eq("hung on the drawing by relationships", [(rels.match(/relationships\/chart"/g) ?? []).length, (rels.match(/2014\/relationships\/chartEx"/g) ?? []).length], [18, 6]);
eq("the sheet names the drawing", /<drawing r:id="rId\d+"\/>/.test(utf8Text(archive.get("xl/worksheets/sheet1.xml")!)), true);
// Every part is XML; each series its cells by reference, the values cached.
for (const n of [...classic, ...extended]) parseXml(utf8Text(archive.get(n)!));
const bar = parseXml(utf8Text(archive.get("xl/charts/chart1.xml")!));
eq("a column chart: its kind, grouping and series' cells", [attr(all(bar, "barDir")[0], "val"), attr(all(bar, "grouping")[0], "val"), all(bar, "ser").length,
  all(bar, "f").map((f) => (f.children[0] as string)).slice(0, 4), attr(all(bar, "ptCount")[1], "val")],
  ["col", "clustered", 3, ["Sheet1!$B$1", "Sheet1!$A$2:$A$6", "Sheet1!$B$2:$B$6", "Sheet1!$C$1"], "5"]);
eq("its title, legend and axis titles", [all(bar, "t").map((t) => t.children[0]), attr(all(bar, "legendPos")[0], "val"), attr(all(bar, "axPos")[0], "val"), attr(all(bar, "axPos")[1], "val")],
  [["bar of sales", "Quarter", "Units"], "t", "b", "l"]);
const waterfall = parseXml(utf8Text(archive.get(`xl/charts/chartEx${kinds.filter((k) => ["histogram", "waterfall", "funnel", "treemap", "sunburst", "box"].includes(k)).indexOf("waterfall") + 1}.xml`)!));
eq("a waterfall: its layout, its categories and values by reference, the Total a subtotal", [attr(all(waterfall, "series")[0], "layoutId"), all(waterfall, "f").map((f) => f.children[0]), attr(all(waterfall, "idx")[0], "val")],
  ["waterfall", ["Sheet1!$A$2:$A$6", "Sheet1!$B$2:$B$6", "Sheet1!$B$1"], "4"]);
const drawing = utf8Text(archive.get("xl/drawings/drawing1.xml")!);
eq("each chart anchored at its cell", [(drawing.match(/<xdr:graphicFrame/g) ?? []).length, drawing.includes("<xdr:col>31</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>")], [24, true]);

// Our own charts, read back as they were.
const wb2 = new Workbook(), cmds2 = new Commands(wb2);
const back = await cmds2.run(`file.importXlsx name=Charts.xlsx base64=${o.output}`);
eq("read back, nothing left out", [back.ok, /left out/.test(back.message)], [true, false]);
const page2 = wb2.page(wb2.active)!;
eq("the same kinds, in order", page2.charts.map((c) => c.kind), sheet.charts.map((c) => c.kind));
eq("their ranges, titles and cells", page2.charts.map((c) => [c.range.map((a) => `${a.col},${a.row}`), c.title, c.at ? `${c.at.col},${c.at.row}` : ""]),
  sheet.charts.map((c) => [c.range.map((a) => `${a.col},${a.row}`), c.title, `${c.at!.col},${c.at!.row}`]));
// Options: the same, but what a kind has no place for (a pie's axis titles
// and gridlines, a treemap's, a stock chart's labels) has nowhere to be kept.
const NO_AXES = ["pie", "doughnut", "treemap", "sunburst"];
const norm = (kind: string, options: Record<string, unknown> | undefined): string => {
  const o: Record<string, unknown> = { ...options };
  if (NO_AXES.includes(kind)) { delete o.xTitle; delete o.yTitle; delete o.gridlines; }
  if (kind === "pie" || kind === "doughnut") delete o.labels;
  if (kind === "stock" || kind === "treemap" || kind === "sunburst") delete o.labels;
  return JSON.stringify(Object.keys(o).sort().map((k) => [k, o[k]]));
};
const differ = sheet.charts.map((c, i) => [c.kind, norm(c.kind, c.options as Record<string, unknown>), norm(c.kind, page2.charts[i].options as Record<string, unknown>)]).filter(([, a, b]) => a !== b);
eq("their options, every one", differ, []);
eq("drawn the same, shape for shape", drawn(wb2).map((d, i) => d === before[i]), before.map(() => true));

// A PivotChart is written as a plain chart, said.
const wbp = new Workbook(), cp = new Commands(wbp);
await cp.run("formulas.use language=Excel");
for (const [k, v] of [["A1", "Region"], ["B1", "Qty"], ["A2", "N"], ["B2", "1"], ["A3", "S"], ["B3", "2"], ["A4", "N"], ["B4", "3"]]) await cp.run(`cell.set cell=${k} input="${v}"`);
await cp.run("chart.insert kind=bar range=A1:B4");
const noted = await cp.run("file.exportXlsx");
eq("a lone chart: nothing said of charts", String(noted.message).includes("chart"), false);

// Excel's own: the parts of a file it wrote, hung on a drawing. The base is
// a workbook with one picture-free drawing, so the sheet, the drawing and
// its relationships are written; the chart parts and their anchors replaced.
const base = new Workbook(), cb = new Commands(base);
await cb.run("formulas.use language=Excel");
for (const [k, v] of [["A1", ""], ["B1", "North"], ["C1", "South"], ["A2", "Q1"], ["A3", "Q2"], ["A4", "Q3"], ["B2", "10"], ["B3", "20"], ["B4", "15"], ["C2", "5"], ["C3", "12"], ["C4", "18"]]) await cb.run(`cell.set cell=${k} input="${v}"`);
await cb.run("sheet.add name=Data");
await cb.run('cell.set cell=A1 input="elsewhere"');
await cb.run("sheet.activate name=Sheet1");
await cb.run("insert.shape shape=rectangle at=H1");
const baseXlsx = unzip(fromBase64(String((await cb.run("file.exportXlsx")).output)));
const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const str = (f: string, v: string[]) => `<c:strRef><c:f>${f}</c:f><c:strCache><c:ptCount val="${v.length}"/>${v.map((x, i) => `<c:pt idx="${i}"><c:v>${x}</c:v></c:pt>`).join("")}</c:strCache></c:strRef>`;
const num = (f: string, v: number[]) => `<c:numRef><c:f>${f}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${v.length}"/>${v.map((x, i) => `<c:pt idx="${i}"><c:v>${x}</c:v></c:pt>`).join("")}</c:numCache></c:numRef>`;
const ser = (i: number, name: string, nf: string, col: string, vs: number[], extra = "") => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${str(nf, [name])}</c:tx><c:spPr><a:solidFill><a:schemeClr val="accent${i + 1}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>${extra}<c:cat>${str("Sheet1!$A$2:$A$4", ["Q1", "Q2", "Q3"])}</c:cat><c:val>${num(col, vs)}</c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{0000000${i}-1234-4567-8901-000000000001}"/></c:ext></c:extLst></c:ser>`;
const title = (t: string) => `<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
const axes = `<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="20"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="9525"/></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="10"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
const wrap = (head: string, plot: string, legend = '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace ${NS} xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${head}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></c:spPr></c:chartSpace>`;
const sers2 = ser(0, "North", "Sheet1!$B$1", "Sheet1!$B$2:$B$4", [10, 20, 15]) + ser(1, "South", "Sheet1!$C$1", "Sheet1!$C$2:$C$4", [5, 12, 18]);
const dl = '<c:dLbls><c:spPr><a:noFill/></c:spPr><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';
const CHARTS: Record<string, string> = {
  column: wrap(title("Quarterly sales"), `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers2}${dl}<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="10"/><c:axId val="20"/></c:barChart>${axes}`),
  column3d: wrap(title("Solid"), `<c:view3D><c:rotX val="15"/></c:view3D><c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers2}<c:gapWidth val="150"/><c:shape val="box"/><c:axId val="10"/><c:axId val="20"/><c:axId val="0"/></c:bar3DChart>${axes}`),
  pctline: wrap(title("Shares"), `<c:lineChart><c:grouping val="percentStacked"/><c:varyColors val="0"/>${sers2}<c:marker val="1"/><c:axId val="10"/><c:axId val="20"/></c:lineChart>${axes}`),
  elsewhere: wrap(title("Another sheet"), `<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$A$1</c:f></c:strRef></c:tx><c:val>${num("Data!$A$1:$A$3", [1, 2, 3])}</c:val></c:ser><c:overlap val="100"/><c:axId val="10"/><c:axId val="20"/></c:barChart>${axes}`),
  literal: wrap(title("Typed in"), `<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart>`),
  pie: wrap("<c:autoTitleDeleted val=\"1\"/>", `<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${str("Sheet1!$B$1", ["North"])}</c:tx><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:cat>${str("Sheet1!$A$2:$A$4", ["Q1", "Q2", "Q3"])}</c:cat><c:val>${num("Sheet1!$B$2:$B$4", [10, 20, 15])}</c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart>`, ""),
  pieBare: wrap("<c:autoTitleDeleted val=\"1\"/>", `<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat>${str("Sheet1!$A$2:$A$4", ["Q1", "Q2", "Q3"])}</c:cat><c:val>${num("Sheet1!$B$2:$B$4", [10, 20, 15])}</c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart>`, ""),
  byRow: wrap(title("By row"), `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${["Q1", "Q2"].map((q, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${str(`Sheet1!$A$${2 + i}`, [q])}</c:tx><c:cat>${str("Sheet1!$B$1:$C$1", ["North", "South"])}</c:cat><c:val>${num(`Sheet1!$B$${2 + i}:$C$${2 + i}`, [10, 5])}</c:val><c:smooth val="0"/></c:ser>`).join("")}<c:marker val="1"/><c:axId val="10"/><c:axId val="20"/></c:lineChart>${axes}`),
};
const EXCHART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">Q1</cx:pt><cx:pt idx="1">Q2</cx:pt><cx:pt idx="2">Q3</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$4</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">10</cx:pt><cx:pt idx="1">20</cx:pt><cx:pt idx="2">15</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>Bridge</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="LAYOUT" uniqueId="{11111111-2222-3333-4444-555555555555}"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>North</cx:v></cx:txData></cx:tx><cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels><cx:dataId val="0"/><cx:layoutPr><cx:subtotals><cx:idx val="2"/></cx:subtotals></cx:layoutPr></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea><cx:legend pos="t" align="ctr" overlay="0"/></cx:chart></cx:chartSpace>`;
const frameOf = (rid: string, id: number, name: string, row: number) => `<xdr:twoCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row + 15}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
const exFrame = (rid: string, id: number, name: string, row: number) => `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row + 15}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
function excel(charts: [string, string, boolean][]): string {
  const parts = new Map(baseXlsx);
  const anchors = charts.map(([_n, _x, ex], i) => (ex ? exFrame : frameOf)(`rId${i + 1}`, i + 2, `Chart ${i + 1}`, i * 16)).join("");
  parts.set("xl/drawings/drawing1.xml", utf8Bytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors}</xdr:wsDr>`));
  parts.set("xl/drawings/_rels/drawing1.xml.rels", utf8Bytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${charts.map(([n, , ex], i) => `<Relationship Id="rId${i + 1}" Type="${ex ? "http://schemas.microsoft.com/office/2014/relationships/chartEx" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"}" Target="../charts/${n}"/>`).join("")}</Relationships>`));
  for (const [n, x] of charts) parts.set(`xl/charts/${n}`, utf8Bytes(x));
  return toBase64(zip([...parts]));
}
const read = (charts: [string, string, boolean][]) => readXlsx(fromBase64(excel(charts)));
const column = read([["chart1.xml", CHARTS.column, false]]);
const first = (column.file.charts as Record<string, unknown>[])[0];
eq("Excel's clustered column chart: its kind, cells, title, legend, place", [first.kind, first.range, first.title, first.at, first.options], ["bar", "A1:C4", "Quarterly sales", "I1", { legend: "right" }]);
eq("nothing said of it", column.report.filter((l) => /chart/.test(l)), []);
const drawn1 = new Workbook(), c1 = new Commands(drawn1);
await c1.run(`file.importXlsx name=Excel.xlsx base64=${excel([["chart1.xml", CHARTS.column, false]])}`);
eq("opened, drawn: two series of three, over the cells", drawn1.charts().map((c) => [c.kind, c.range, c.data.labels, c.data.series.map((s) => s.values)]), [["bar", "A1:C4", ["Q1", "Q2", "Q3"], [[10, 20, 15], [5, 12, 18]]]]);
const hard = read([["chart1.xml", CHARTS.column3d, false], ["chart2.xml", CHARTS.pctline, false], ["chart3.xml", CHARTS.elsewhere, false], ["chart4.xml", CHARTS.literal, false]]);
eq("a 3-D chart, a 100% stacked line, cells of another sheet, data typed in: each said, none read", [hard.report.filter((l) => /chart/.test(l)), (hard.file.charts as unknown[] | undefined)?.length ?? 0],
  [["Sheet1: the chart Solid left out: a 3-D column or bar chart, which the app does not draw", "Sheet1: the chart Shares left out: a 100% stacked line chart, which the app does not have",
    "Sheet1: the chart Another sheet left out: it reads cells of Data", "Sheet1: the chart Typed in left out: it holds some of its data itself, not cells"], 0]);
const mixed = read([["chart1.xml", CHARTS.column, false], ["chart2.xml", CHARTS.column3d, false], ["chart3.xml", CHARTS.pie, false], ["chart4.xml", CHARTS.byRow, false], ["chart5.xml", CHARTS.pieBare, false]]);
const got = mixed.file.charts as Record<string, unknown>[];
eq("a file with some of each: the modelled read, the others said once", [got.map((c) => [c.kind, c.range, c.at]), mixed.report.filter((l) => /chart/.test(l)).length],
  [[["bar", "A1:C4", "I1"], ["pie", "A1:B4", "I33"], ["line", "A1:C3", "I49"], ["pie", "A2:B4", "I65"]], 1]);
eq("a pie showing percents keeps the labels it has; one showing none has them off; no legend is none; the rows the series", [got[1].options, got[3].options, got[2].options], [{ legend: "none" }, { legend: "none", labels: false }, { legend: "right", byRows: true }]);
for (const [layout, kind] of [["waterfall", "waterfall"], ["funnel", "funnel"], ["treemap", "treemap"], ["sunburst", "sunburst"], ["boxWhisker", "box"], ["regionMap", ""]] as const) {
  const r = read([["chartEx1.xml", EXCHART.replace("LAYOUT", layout), true]]);
  const c = (r.file.charts as Record<string, unknown>[] | undefined)?.[0];
  eq(`Excel's ${layout} chart part`, c ? [c.kind, c.range, c.title, c.options] : r.report.filter((l) => /chart/.test(l)),
    kind ? [kind, "A1:B4", "Bridge", { legend: "top", labels: true }] : ["Sheet1: the chart Bridge left out: a map of a kind the app does not have"]);
}
const histo = read([["chartEx1.xml", EXCHART.replace("LAYOUT", "clusteredColumn").replace("<cx:subtotals><cx:idx val=\"2\"/></cx:subtotals>", '<cx:binning intervalClosed="r"/>'), true]]);
eq("a histogram is the clustered column with its bins", (histo.file.charts as Record<string, unknown>[])[0].kind, "histogram");
eq("a reference's parts", [parseRef("'My Sheet'!$B$2:$D$9"), parseRef("Sheet1!A1"), parseRef("(Sheet1!A1,Sheet1!B2)"), parseRef("Sheet1!$A:$A")],
  [{ sheet: "My Sheet", r: { a: { col: 1, row: 1 }, b: { col: 3, row: 8 } } }, { sheet: "Sheet1", r: { a: { col: 0, row: 0 }, b: { col: 0, row: 0 } } }, null, null]);

// expect: ok the charts made, one of each kind
// expect: ok a part for each: eighteen classic, six extended
// expect: ok said in the report: not that the charts are left out
// expect: ok the parts' types
// expect: ok hung on the drawing by relationships
// expect: ok the sheet names the drawing
// expect: ok a column chart: its kind, grouping and series' cells
// expect: ok its title, legend and axis titles
// expect: ok a waterfall: its layout, its categories and values by reference, the Total a subtotal
// expect: ok each chart anchored at its cell
// expect: ok read back, nothing left out
// expect: ok the same kinds, in order
// expect: ok their ranges, titles and cells
// expect: ok their options, every one
// expect: ok drawn the same, shape for shape
// expect: ok a lone chart: nothing said of charts
// expect: ok Excel's clustered column chart: its kind, cells, title, legend, place
// expect: ok nothing said of it
// expect: ok opened, drawn: two series of three, over the cells
// expect: ok a 3-D chart, a 100% stacked line, cells of another sheet, data typed in: each said, none read
// expect: ok a file with some of each: the modelled read, the others said once
// expect: ok a pie showing percents keeps the labels it has; one showing none has them off; no legend is none; the rows the series
// expect: ok Excel's waterfall chart part
// expect: ok Excel's funnel chart part
// expect: ok Excel's treemap chart part
// expect: ok Excel's sunburst chart part
// expect: ok Excel's boxWhisker chart part
// expect: ok Excel's regionMap chart part
// expect: ok a histogram is the clustered column with its bins
// expect: ok a reference's parts
