// A sheet's charts in an XLSX file: DrawingML chart parts (ECMA-376 21.2),
// written from the app's charts and read back. The classic parts (c:chartSpace)
// hold columns, bars, lines, areas, pies, doughnuts, scatter, bubble, radar,
// stock and the combination; the extended ones (cx:chartSpace, Office 2016's)
// the histogram, waterfall, funnel, treemap, sunburst and box and whisker.
// A chart is its series as references to the sheet's cells (with the values
// cached beside them, as Excel writes them), its title, legend, axis titles,
// data labels, gridlines and trendlines; read back, the cells the series name
// are the chart's range, so our own charts come back as they were. What the
// app draws its own way (Excel's colours, fonts, 3-D) is not kept; a chart of
// a kind the app does not have, or reading cells that are not one block on
// its own sheet, is said in the report and left out.
import { attr, all, escapeXml, kid, kids, local, textOf } from "./xml.ts";
import type { El } from "./xml.ts";
import { addrOf, colName, keyOf } from "./address.ts";
import type { Addr } from "./address.ts";
import { chartData, chartShape } from "./charts.ts";
import type { ChartData, ChartKind, ChartOptions, Trendline } from "./charts.ts";
import type { Value } from "./values.ts";
import { TREND_TYPES } from "./trend.ts";
import type { TrendType } from "./trend.ts";

export const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const CX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// The kinds Office 2016's extended chart part holds.
export const EX_KINDS: readonly ChartKind[] = ["histogram", "waterfall", "funnel", "treemap", "sunburst", "box"];
export const isEx = (kind: ChartKind): boolean => EX_KINDS.includes(kind);

// ---- references -----------------------------------------------------------

interface Rect { readonly a: Addr; readonly b: Addr }
const bbox = (cells: readonly Addr[]): Rect => ({
  a: { col: Math.min(...cells.map((c) => c.col)), row: Math.min(...cells.map((c) => c.row)) },
  b: { col: Math.max(...cells.map((c) => c.col)), row: Math.max(...cells.map((c) => c.row)) },
});
// A sheet's name as a reference writes it: quoted unless a plain word.
const quoted = (sheet: string): string => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`);
const absolute = (a: Addr): string => `$${colName(a.col)}$${a.row + 1}`;
const refText = (sheet: string, r: Rect): string =>
  `${quoted(sheet)}!${absolute(r.a)}${r.a.col === r.b.col && r.a.row === r.b.row ? "" : `:${absolute(r.b)}`}`;

// "Sheet1!$A$2:$A$6" as its sheet and cells; null for a union, a name, a
// whole column or anything else that is not one block of cells.
export function parseRef(f: string): { sheet: string; r: Rect } | null {
  const m = /^\s*(?:'((?:[^']|'')+)'|([^'!()]+))!\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?\s*$/.exec(f);
  if (!m) return null;
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const a = addrOf(`${m[3]}${m[4]}`), b = addrOf(`${m[5] ?? m[3]}${m[6] ?? m[4]}`);
  return a && b ? { sheet, r: bbox([a, b]) } : null;
}

// ---- what a chart is made of ----------------------------------------------

interface Ser {
  readonly name: string;
  readonly nameRef?: string;
  readonly ref: string;
  readonly values: readonly number[];
}
interface Plan {
  readonly data: ChartData;
  readonly sers: readonly Ser[];
  // The categories shared by the series (the first column or two).
  readonly cat?: { readonly ref: string; readonly labels: readonly string[]; readonly groups?: readonly string[] };
}

export interface ChartIn {
  readonly kind: ChartKind;
  readonly range: [Addr, Addr];
  readonly title: string;
  readonly options?: ChartOptions;
  readonly at: Addr;
  // A PivotChart's cells: what its pivot last wrote.
  readonly written?: [Addr, Addr] | null;
  readonly pivot?: boolean;
}

// The data of a chart and where each series is: its name's cell, its values,
// its categories, as Excel reads the range (charts.ts's chartShape).
function plan(c: ChartIn, host: string, value: (key: string) => Value): Plan {
  const corners = c.written ?? c.range, byRows = !!c.options?.byRows;
  const shape = chartShape(corners, value, byRows);
  const data = chartData(corners, value, byRows);
  const lo = { row: Math.min(corners[0].row, corners[1].row), col: Math.min(corners[0].col, corners[1].col) };
  const at = (i: number, j: number): Addr => (byRows ? { row: lo.row + j, col: lo.col + i } : { row: lo.row + i, col: lo.col + j });
  const top = shape.headed ? 1 : 0, last = shape.grid.length - 1;
  // A PivotChart leaves its grand totals out.
  const rows: number[] = [];
  for (let i = top; i <= last; i++) if (!c.pivot || data.labels[i - top] !== "Grand Total") rows.push(i);
  const keep = c.pivot ? data.series.map((s, j) => (s.name === "Grand Total" ? -1 : j)).filter((j) => j >= 0) : data.series.map((_, j) => j);
  const labelsOf = rows.map((i) => data.labels[i - top]);
  const sers: Ser[] = keep.map((j) => ({
    name: data.series[j].name,
    ...(shape.headed ? { nameRef: refText(host, bbox([at(0, shape.first + j)])) } : {}),
    ref: refText(host, bbox(rows.map((i) => at(i, shape.first + j)))),
    values: rows.map((i) => data.series[j].values[i - top]),
  }));
  const cat = shape.labelled && rows.length
    ? { ref: refText(host, bbox(rows.flatMap((i) => (shape.twoLevel ? [at(i, 0), at(i, 1)] : [at(i, 0)])))), labels: labelsOf,
        ...(shape.twoLevel ? { groups: rows.map((i) => (data.groups ?? [])[i - top] ?? "") } : {}) }
    : undefined;
  return { data: { ...data, labels: labelsOf, series: keep.map((j, k) => ({ ...data.series[j], values: sers[k].values })) }, sers, cat };
}

// ---- XML pieces -----------------------------------------------------------

const val = (name: string, v: string | number): string => `<c:${name} val="${v}"/>`;
const pts = (items: readonly (string | number)[]): string => items.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(String(v))}</c:v></c:pt>`).join("");
const strRef = (f: string, items: readonly string[]): string => `<c:strRef><c:f>${escapeXml(f)}</c:f><c:strCache><c:ptCount val="${items.length}"/>${pts(items)}</c:strCache></c:strRef>`;
const numRef = (f: string, items: readonly number[]): string => `<c:numRef><c:f>${escapeXml(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${items.length}"/>${pts(items)}</c:numCache></c:numRef>`;
const rich = (text: string): string => `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
const titleXml = (text: string | undefined): string => (text ? `<c:title>${rich(text)}<c:overlay val="0"/></c:title>` : "");

function catXml(p: Plan): string {
  if (!p.cat) return "";
  if (!p.cat.groups) return `<c:cat>${strRef(p.cat.ref, p.cat.labels)}</c:cat>`;
  return `<c:cat><c:multiLvlStrRef><c:f>${escapeXml(p.cat.ref)}</c:f><c:multiLvlStrCache><c:ptCount val="${p.cat.labels.length}"/><c:lvl>${pts(p.cat.labels)}</c:lvl><c:lvl>${pts(p.cat.groups)}</c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`;
}
const nameXml = (s: Ser): string => `<c:tx>${s.nameRef ? strRef(s.nameRef, [s.name]) : `<c:v>${escapeXml(s.name)}</c:v>`}</c:tx>`;
const TREND: Readonly<Record<TrendType, string>> = { linear: "linear", exponential: "exp", logarithmic: "log", polynomial: "poly", power: "power", movingAverage: "movingAvg" };
function trendXml(o: ChartOptions, j: number): string {
  return (o.trendlines ?? []).filter((t) => t.series === j).map((t) =>
    `<c:trendline><c:trendlineType val="${TREND[t.type]}"/>${t.type === "polynomial" ? val("order", Math.max(2, Math.min(6, t.order ?? 2))) : ""}${t.type === "movingAverage" ? val("period", Math.max(2, t.period ?? 2)) : ""}${val("dispRSqr", t.r2 ? 1 : 0)}${val("dispEq", t.equation ? 1 : 0)}</c:trendline>`).join("");
}
const MARKER = '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>';
const dLbls = (percent: boolean): string => `<c:dLbls>${val("showLegendKey", 0)}${val("showVal", percent ? 0 : 1)}${val("showCatName", 0)}${val("showSerName", 0)}${val("showPercent", percent ? 1 : 0)}${val("showBubbleSize", 0)}</c:dLbls>`;

// The axes: a pair of ids, the category (or x) axis first.
function axisXml(kind: "cat" | "val", id: number, cross: number, pos: string, o: { title?: string; grid?: boolean; hidden?: boolean; max?: boolean; percent?: boolean; between?: boolean }): string {
  const scaling = "<c:scaling><c:orientation val=\"minMax\"/></c:scaling>";
  const common = `${val("axId", id)}${scaling}${val("delete", o.hidden ? 1 : 0)}${val("axPos", pos)}`;
  if (kind === "cat") {
    return `<c:catAx>${common}${titleXml(o.title)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${val("crossAx", cross)}<c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  }
  return `<c:valAx>${common}${o.grid ? "<c:majorGridlines/>" : ""}${titleXml(o.title)}<c:numFmt formatCode="${o.percent ? "0%" : "General"}" sourceLinked="${o.percent ? 0 : 1}"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${val("crossAx", cross)}<c:crosses val="${o.max ? "max" : "autoZero"}"/>${val("crossBetween", o.between === false ? "midCat" : "between")}</c:valAx>`;
}

// A chart as a classic chart part.
export function chartXml(c: ChartIn, host: string, value: (key: string) => Value): string {
  const p = plan(c, host, value), o = c.options ?? {}, k = c.kind, n = p.sers.length;
  const grid = o.gridlines !== false;
  const sers = (each: (s: Ser, j: number) => string, from = 0, to = n): string => p.sers.slice(from, to).map((s, i) => `<c:ser>${val("idx", from + i)}${val("order", from + i)}${each(s, from + i)}</c:ser>`).join("");
  const plain = (s: Ser, j: number, extra = ""): string => `${nameXml(s)}${extra}${trendXml(o, j)}${catXml(p)}<c:val>${numRef(s.ref, s.values)}</c:val>`;
  let groups = "", axes = "";
  let hTitle = o.xTitle, vTitle = o.yTitle;   // the horizontal axis's, the vertical's
  if (k === "bar" || k === "barStacked" || k === "bar100" || k === "hbar" || k === "hbarStacked" || k === "hbar100") {
    const horizontal = k.startsWith("hbar"), grouping = k.endsWith("100") ? "percentStacked" : /Stacked$/.test(k) ? "stacked" : "clustered";
    groups = `<c:barChart>${val("barDir", horizontal ? "bar" : "col")}${val("grouping", grouping)}${val("varyColors", 0)}${sers((s, j) => plain(s, j, "<c:invertIfNegative val=\"0\"/>"))}${o.labels ? dLbls(false) : ""}${val("gapWidth", 150)}${grouping === "clustered" ? "" : val("overlap", 100)}${val("axId", 111)}${val("axId", 112)}</c:barChart>`;
    axes = axisXml("cat", 111, 112, horizontal ? "l" : "b", { title: horizontal ? vTitle : hTitle }) + axisXml("val", 112, 111, horizontal ? "b" : "l", { title: horizontal ? hTitle : vTitle, grid, percent: grouping === "percentStacked" });
  } else if (k === "line" || k === "lineStacked") {
    groups = `<c:lineChart>${val("grouping", k === "line" ? "standard" : "stacked")}${val("varyColors", 0)}${sers((s, j) => `${plain(s, j, MARKER)}<c:smooth val="0"/>`)}${o.labels ? dLbls(false) : ""}${val("marker", 1)}${val("axId", 111)}${val("axId", 112)}</c:lineChart>`;
    axes = axisXml("cat", 111, 112, "b", { title: hTitle }) + axisXml("val", 112, 111, "l", { title: vTitle, grid });
  } else if (k === "area" || k === "areaStacked") {
    groups = `<c:areaChart>${val("grouping", k === "area" ? "standard" : "stacked")}${val("varyColors", 0)}${sers((s, j) => plain(s, j))}${o.labels ? dLbls(false) : ""}${val("axId", 111)}${val("axId", 112)}</c:areaChart>`;
    axes = axisXml("cat", 111, 112, "b", { title: hTitle }) + axisXml("val", 112, 111, "l", { title: vTitle, grid });
  } else if (k === "pie" || k === "doughnut") {
    const one = `<c:ser>${val("idx", 0)}${val("order", 0)}${p.sers[0] ? `${nameXml(p.sers[0])}${o.labels === false ? "" : dLbls(true)}${catXml(p)}<c:val>${numRef(p.sers[0].ref, p.sers[0].values)}</c:val>` : ""}</c:ser>`;
    groups = k === "pie" ? `<c:pieChart>${val("varyColors", 1)}${one}${val("firstSliceAng", 0)}</c:pieChart>` : `<c:doughnutChart>${val("varyColors", 1)}${one}${val("firstSliceAng", 0)}${val("holeSize", 50)}</c:doughnutChart>`;
  } else if (k === "radar") {
    groups = `<c:radarChart>${val("radarStyle", "marker")}${val("varyColors", 0)}${sers((s, j) => `${nameXml(s)}${MARKER}${catXml(p)}<c:val>${numRef(s.ref, s.values)}</c:val>`)}${o.labels ? dLbls(false) : ""}${val("axId", 111)}${val("axId", 112)}</c:radarChart>`;
    axes = axisXml("cat", 111, 112, "b", { title: hTitle }) + axisXml("val", 112, 111, "l", { title: vTitle, grid });
  } else if (k === "scatter" || k === "scatterLines" || k === "bubble") {
    // The first column is every series' x where it is numbers (the layout's
    // "counted"); else the labels are the x.
    const counted = !p.cat && n > 1;
    const x = counted ? `<c:xVal>${numRef(p.sers[0].ref, p.sers[0].values)}</c:xVal>` : p.cat ? `<c:xVal>${strRef(p.cat.ref, p.cat.labels)}</c:xVal>` : "";
    const ys = counted ? p.sers.slice(1) : p.sers;
    const line = k === "scatterLines" ? '<c:spPr><a:ln w="28575" cap="rnd"><a:round/></a:ln></c:spPr>' : '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>';
    let body: string;
    if (k === "bubble") {
      body = ys.filter((_, i) => i % 2 === 0).map((s, i) => {
        const size = ys[i * 2 + 1];
        return `<c:ser>${val("idx", i)}${val("order", i)}${nameXml(s)}<c:invertIfNegative val="0"/>${x}<c:yVal>${numRef(s.ref, s.values)}</c:yVal>${size ? `<c:bubbleSize>${numRef(size.ref, size.values)}</c:bubbleSize>` : ""}<c:bubble3D val="0"/></c:ser>`;
      }).join("");
      groups = `<c:bubbleChart>${val("varyColors", 0)}${body}${o.labels ? dLbls(false) : ""}${val("bubbleScale", 100)}${val("showNegBubbles", 0)}${val("axId", 111)}${val("axId", 112)}</c:bubbleChart>`;
    } else {
      body = ys.map((s, i) => `<c:ser>${val("idx", i)}${val("order", i)}${nameXml(s)}${line}${MARKER}${trendXml(o, i)}${x}<c:yVal>${numRef(s.ref, s.values)}</c:yVal><c:smooth val="0"/></c:ser>`).join("");
      groups = `<c:scatterChart>${val("scatterStyle", "lineMarker")}${val("varyColors", 0)}${body}${o.labels ? dLbls(false) : ""}${val("axId", 111)}${val("axId", 112)}</c:scatterChart>`;
    }
    axes = axisXml("val", 111, 112, "b", { title: hTitle, between: false }) + axisXml("val", 112, 111, "l", { title: vTitle, grid, between: false });
  } else if (k === "stock") {
    const four = n >= 4;
    groups = `<c:stockChart>${sers((s, j) => `${nameXml(s)}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="${!four && j === n - 1 ? "dash" : "none"}"/><c:size val="5"/></c:marker>${catXml(p)}<c:val>${numRef(s.ref, s.values)}</c:val><c:smooth val="0"/>`)}<c:hiLowLines/>${four ? "<c:upDownBars><c:gapWidth val=\"150\"/><c:upBars/><c:downBars/></c:upDownBars>" : ""}${val("axId", 111)}${val("axId", 112)}</c:stockChart>`;
    axes = axisXml("cat", 111, 112, "b", { title: hTitle }) + axisXml("val", 112, 111, "l", { title: vTitle, grid });
  } else if (k === "combo") {
    // The first series columns, the rest lines; those asked for on the
    // secondary axes (ids 113 and 114) with the right axis showing.
    const second = new Set(o.secondary ?? []);
    const axisIds = (j: number) => (second.has(j) ? [113, 114] : [111, 112]);
    const bar = n ? `<c:barChart>${val("barDir", "col")}${val("grouping", "clustered")}${val("varyColors", 0)}${sers((s, j) => plain(s, j, "<c:invertIfNegative val=\"0\"/>"), 0, 1)}${o.labels ? dLbls(false) : ""}${val("gapWidth", 150)}${axisIds(0).map((i) => val("axId", i)).join("")}</c:barChart>` : "";
    const lineGroup = (secondary: boolean): string => {
      const idx = p.sers.map((_, j) => j).filter((j) => j > 0 && second.has(j) === secondary);
      if (!idx.length) return "";
      const ids = secondary ? [113, 114] : [111, 112];
      return `<c:lineChart>${val("grouping", "standard")}${val("varyColors", 0)}${idx.map((j) => `<c:ser>${val("idx", j)}${val("order", j)}${plain(p.sers[j], j, MARKER)}<c:smooth val="0"/></c:ser>`).join("")}${o.labels ? dLbls(false) : ""}${val("marker", 1)}${ids.map((i) => val("axId", i)).join("")}</c:lineChart>`;
    };
    groups = bar + lineGroup(false) + lineGroup(true);
    const usesPrimary = p.sers.some((_, j) => !second.has(j)), usesSecond = p.sers.some((_, j) => second.has(j));
    axes = (usesPrimary ? axisXml("cat", 111, 112, "b", { title: hTitle }) + axisXml("val", 112, 111, "l", { title: vTitle, grid }) : "")
      + (usesSecond ? axisXml("cat", 113, 114, "b", { hidden: true }) + axisXml("val", 114, 113, "r", { max: true }) : "");
  }
  const legend = (o.legend ?? "bottom") === "none" ? "" : `<c:legend>${val("legendPos", ({ bottom: "b", top: "t", right: "r", left: "l" } as Record<string, string>)[o.legend ?? "bottom"] ?? "b")}<c:overlay val="0"/></c:legend>`;
  return `${HEAD}<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"><c:roundedCorners val="0"/><c:chart>${titleXml(c.title)}${val("autoTitleDeleted", c.title ? 0 : 1)}<c:plotArea><c:layout/>${groups}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}

// A chart as an extended chart part (Office 2016's): its data by reference,
// then its one series (a box: one for each column).
export function chartExXml(c: ChartIn, host: string, value: (key: string) => Value): string {
  const p = plan(c, host, value), o = c.options ?? {}, k = c.kind;
  const lvl = (items: readonly (string | number)[]) => `<cx:lvl ptCount="${items.length}">${items.map((v, i) => `<cx:pt idx="${i}">${escapeXml(String(v))}</cx:pt>`).join("")}</cx:lvl>`;
  const cat = p.cat
    ? `<cx:strDim type="cat"><cx:f>${escapeXml(p.cat.ref)}</cx:f>${p.cat.groups ? `${lvl(p.cat.labels)}${lvl(p.cat.groups)}` : lvl(p.cat.labels)}</cx:strDim>` : "";
  const num = (s: Ser) => `<cx:numDim type="val"><cx:f>${escapeXml(s.ref)}</cx:f><cx:lvl ptCount="${s.values.length}" formatCode="General">${s.values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join("")}</cx:lvl></cx:numDim>`;
  const used = k === "box" ? p.sers : p.sers.slice(0, 1);
  const data = used.map((s, i) => `<cx:data id="${i}">${cat}${num(s)}</cx:data>`).join("");
  const layoutId = ({ histogram: "clusteredColumn", waterfall: "waterfall", funnel: "funnel", treemap: "treemap", sunburst: "sunburst", box: "boxWhisker" } as Record<string, string>)[k];
  const axisIds = k === "treemap" || k === "sunburst" ? "" : '<cx:axisId val="0"/><cx:axisId val="1"/>';
  const layoutPr = k === "histogram" ? '<cx:layoutPr><cx:binning intervalClosed="r"/></cx:layoutPr>'
    : k === "waterfall" ? `<cx:layoutPr><cx:subtotals>${(p.cat?.labels ?? []).map((l, i) => (/^(sub)?total$/i.test(l.trim()) ? `<cx:idx val="${i}"/>` : "")).join("")}</cx:subtotals></cx:layoutPr>`
    : k === "box" ? '<cx:layoutPr><cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/><cx:statistics quartileMethod="exclusive"/></cx:layoutPr>'
    : k === "treemap" ? `<cx:layoutPr><cx:parentLabelLayout val="${p.cat?.groups ? "banner" : "none"}"/></cx:layoutPr>` : "";
  const labels = o.labels ? '<cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>' : "";
  const series = used.map((s, i) => `<cx:series layoutId="${layoutId}" uniqueId="{00000000-0000-4000-8000-${String(i + 1).padStart(12, "0")}}"><cx:tx><cx:txData>${s.nameRef ? `<cx:f>${escapeXml(s.nameRef)}</cx:f>` : ""}<cx:v>${escapeXml(s.name)}</cx:v></cx:txData></cx:tx>${labels}<cx:dataId val="${i}"/>${layoutPr}${axisIds}</cx:series>`).join("");
  const cxTitle = (t: string | undefined) => (t ? `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>${escapeXml(t)}</cx:v></cx:txData></cx:tx></cx:title>` : "");
  const axes = axisIds
    ? `<cx:axis id="0">${k === "histogram" ? '<cx:catScaling gapWidth="0"/>' : '<cx:catScaling gapWidth="0.5"/>'}${cxTitle(o.xTitle)}<cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/>${cxTitle(o.yTitle)}${o.gridlines === false ? "" : "<cx:majorGridlines/>"}<cx:tickLabels/></cx:axis>` : "";
  const legend = (o.legend ?? "bottom") === "none" ? "" : `<cx:legend pos="${({ bottom: "b", top: "t", right: "r", left: "l" } as Record<string, string>)[o.legend ?? "bottom"] ?? "b"}" align="ctr" overlay="0"/>`;
  return `${HEAD}<cx:chartSpace xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:cx="${CX_NS}"><cx:chartData>${data}</cx:chartData><cx:chart>${cxTitle(c.title)}<cx:plotArea><cx:plotAreaRegion>${series}</cx:plotAreaRegion>${axes}</cx:plotArea>${legend}</cx:chart></cx:chartSpace>`;
}

// ---- reading --------------------------------------------------------------

export interface ChartOut {
  readonly kind: ChartKind;
  readonly range: string;
  readonly title: string;
  readonly options?: ChartOptions;
  readonly at: string;
}

// The cells a chart's series name, as one block on the chart's own sheet: or
// why not.
function blockOf(refs: readonly string[], host: string): { r: Rect } | string {
  const parsed = refs.map(parseRef);
  if (!parsed.length) return "it holds its data itself, not cells";
  if (parsed.some((x) => !x)) return "it reads cells that are not one block";
  const sheets = new Set(parsed.map((x) => x!.sheet.toLowerCase()));
  if (sheets.size > 1 || !sheets.has(host.toLowerCase())) return `it reads cells of ${sheets.size > 1 ? "several sheets" : [...parsed][0]!.sheet}`;
  return { r: bbox(parsed.flatMap((x) => [x!.r.a, x!.r.b])) };
}

// A reference's formula, from a c:strRef, c:numRef, c:multiLvlStrRef or the
// extended part's cx:f.
const fOf = (e: El | undefined): string | undefined => {
  if (!e) return undefined;
  const f = all(e, "f")[0];
  return f ? textOf(f).trim() : undefined;
};
const flag = (e: El | undefined, name: string, dflt: boolean): boolean => {
  const x = e && all(e, name)[0];
  if (!x) return dflt;
  const v = attr(x, "val");
  return v === undefined ? true : v === "1" || v === "true";
};

const BAR_KIND: Readonly<Record<string, ChartKind>> = { "col/clustered": "bar", "col/stacked": "barStacked", "col/percentStacked": "bar100", "bar/clustered": "hbar", "bar/stacked": "hbarStacked", "bar/percentStacked": "hbar100" };
const GROUPS = ["barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "radarChart", "scatterChart", "bubbleChart", "stockChart"];
const WORDS: Readonly<Record<string, string>> = {
  bar3DChart: "3-D column or bar", line3DChart: "3-D line", area3DChart: "3-D area", pie3DChart: "3-D pie", ofPieChart: "pie-of-pie or bar-of-pie", surfaceChart: "surface", surface3DChart: "3-D surface",
};

// One chart part read (a c:chartSpace or a cx:chartSpace), as the sheet
// file's chart, or the reason it cannot be. `at` is where its anchor puts it.
export function readChart(part: El, host: string, at: Addr): ChartOut | string {
  return kid(part, "chartData") ? readEx(part, host, at) : readClassic(part, host, at);
}

// A title's text: the runs of a classic chart's rich text, or the extended
// chart's own value.
const titleOf = (e: El | undefined): string => (e ? (all(e, "t").length ? all(e, "t") : all(e, "v")).map(textOf).join("").trim() : "");

function readClassic(part: El, host: string, at: Addr): ChartOut | string {
  const chart = kid(part, "chart"), plot = chart && kid(chart, "plotArea");
  if (!chart || !plot) return "it has no plot area";
  const groups = kids(plot).filter((g) => /Chart$/.test(local(g.name)));
  if (!groups.length) return "it has no chart in its plot area";
  const names = groups.map((g) => local(g.name));
  const odd = names.find((n) => !GROUPS.includes(n));
  if (odd) return `a ${WORDS[odd] ?? odd.replace(/Chart$/, "")} chart, which the app does not draw`;
  // The series of every group, by their order in the chart.
  const sers = groups.flatMap((g) => kids(g, "ser").map((s) => ({ g, s }))).sort((x, y) => Number(attr(kid(x.s, "order") ?? x.s, "val") ?? 0) - Number(attr(kid(y.s, "order") ?? y.s, "val") ?? 0));
  if (!sers.length) return "it has no series";
  // Axes: which are the secondary (a value axis crossing at the maximum).
  const axesOf = (g: El) => kids(g, "axId").map((a) => attr(a, "val") ?? "");
  const valAxes = kids(plot, "valAx"), catAxes = kids(plot, "catAx");
  const secondaryIds = new Set(valAxes.filter((v) => attr(kid(v, "crosses") ?? v, "val") === "max").map((v) => attr(kid(v, "axId") ?? v, "val") ?? ""));
  const isSecondary = (g: El) => axesOf(g).some((id) => secondaryIds.has(id));
  let kind: ChartKind;
  const g0 = groups[0], n0 = names[0];
  let scatterLines = false;
  if (groups.length > 1) {
    // A column series with lines: the combination.
    const bars = groups.filter((g) => local(g.name) === "barChart"), lines = groups.filter((g) => local(g.name) === "lineChart");
    if (groups.length !== bars.length + lines.length || bars.length !== 1 || kids(bars[0], "ser").length !== 1 || !lines.length) return "a combination of charts that is not columns with lines";
    if (sers[0].g !== bars[0] || attr(kid(bars[0], "barDir") ?? bars[0], "val") === "bar") return "a combination that does not begin with a column series";
    kind = "combo";
  } else if (n0 === "barChart") {
    const k = BAR_KIND[`${attr(kid(g0, "barDir") ?? g0, "val") ?? "col"}/${attr(kid(g0, "grouping") ?? g0, "val") ?? "clustered"}`];
    if (!k) return "a bar chart of a kind the app does not have";
    kind = k;
  } else if (n0 === "lineChart" || n0 === "areaChart") {
    const grouping = attr(kid(g0, "grouping") ?? g0, "val") ?? "standard";
    if (grouping === "percentStacked") return `a 100% stacked ${n0 === "lineChart" ? "line" : "area"} chart, which the app does not have`;
    kind = n0 === "lineChart" ? (grouping === "stacked" ? "lineStacked" : "line") : grouping === "stacked" ? "areaStacked" : "area";
  } else if (n0 === "pieChart") kind = "pie";
  else if (n0 === "doughnutChart") kind = "doughnut";
  else if (n0 === "radarChart") kind = "radar";
  else if (n0 === "bubbleChart") kind = "bubble";
  else if (n0 === "stockChart") kind = "stock";
  else {
    // Scatter: lines where a series has its line drawn.
    scatterLines = sers.some(({ s }) => { const sp = kid(s, "spPr"), ln = sp && all(sp, "ln")[0]; return !(ln && kid(ln, "noFill")); });
    kind = scatterLines ? "scatterLines" : "scatter";
  }
  // The cells it reads.
  const refs: string[] = [];
  let missing = false;
  const take = (e: El | undefined, what: string): string | undefined => {
    if (!e) return undefined;
    const f = fOf(e);
    if (!f) { if (what !== "tx") missing = true; return undefined; }
    refs.push(f);
    return f;
  };
  const onlyFirst = kind === "pie" || kind === "doughnut";
  const used = onlyFirst ? sers.slice(0, 1) : sers;
  used.forEach(({ s }) => {
    take(kid(s, "tx"), "tx");
    take(kid(s, "cat"), "cat");
    take(kid(s, "xVal"), "cat");
    if (!kid(s, "val") && !kid(s, "yVal")) missing = true;
    for (const v of ["val", "yVal", "bubbleSize"]) take(kid(s, v), "val");
  });
  if (missing) return "it holds some of its data itself, not cells";
  const block = blockOf(refs, host);
  if (typeof block === "string") return block;
  const firstVal = parseRef(fOf(kid(used[0].s, "val") ?? kid(used[0].s, "yVal") ?? used[0].s) ?? "");
  const byRows = !!firstVal && firstVal.r.a.row === firstVal.r.b.row && firstVal.r.a.col !== firstVal.r.b.col;
  const options: Record<string, unknown> = {};
  const title = titleOf(kid(chart, "title"));
  const legendEl = kid(chart, "legend");
  const pos = legendEl ? attr(kid(legendEl, "legendPos") ?? legendEl, "val") ?? "r" : "";
  if (!legendEl) options.legend = "none";
  else if (pos === "t") options.legend = "top";
  else if (pos === "r" || pos === "tr") options.legend = "right";
  else if (pos === "l") options.legend = "left";
  // Data labels: any series' or group's showing values (a pie's, percents).
  const labelEls = [...groups.flatMap((g) => kids(g, "dLbls")), ...used.flatMap(({ s }) => kids(s, "dLbls"))];
  const shown = labelEls.some((d) => flag(d, "showVal", false) || flag(d, "showPercent", false));
  if (onlyFirst) { if (!shown) options.labels = false; } else if (shown) options.labels = true;
  // Axes: gridlines of the value axis, titles by where the axis lies.
  const primaryVal = valAxes.find((v) => !secondaryIds.has(attr(kid(v, "axId") ?? v, "val") ?? "")) ?? valAxes[0];
  if (primaryVal && !onlyFirst) {
    const valued = kind === "scatter" || kind === "scatterLines" || kind === "bubble" ? valAxes.filter((v) => !secondaryIds.has(attr(kid(v, "axId") ?? v, "val") ?? "")) : [primaryVal];
    const yAxis = valued.find((v) => ["l", "r"].includes(attr(kid(v, "axPos") ?? v, "val") ?? "")) ?? primaryVal;
    if (!kid(yAxis, "majorGridlines")) options.gridlines = false;
  }
  for (const ax of [...catAxes, ...valAxes].filter((a) => !secondaryIds.has(attr(kid(a, "axId") ?? a, "val") ?? ""))) {
    const t = titleOf(kid(ax, "title"));
    if (!t) continue;
    const where = attr(kid(ax, "axPos") ?? ax, "val") ?? "b";
    if (where === "b" || where === "t") options.xTitle = t;
    else options.yTitle = t;
  }
  if (byRows) options.byRows = true;
  // Trendlines, by their series' place in the chart.
  const trendlines: Trendline[] = [];
  used.forEach(({ s }, j) => {
    for (const t of kids(s, "trendline")) {
      const type = (Object.entries(TREND) as [TrendType, string][]).find(([, v]) => v === attr(kid(t, "trendlineType") ?? t, "val"))?.[0];
      if (!type || !(TREND_TYPES as readonly string[]).includes(type)) continue;
      trendlines.push({ series: j, type, ...(type === "polynomial" ? { order: Number(attr(kid(t, "order") ?? t, "val") ?? 2) } : {}), ...(type === "movingAverage" ? { period: Number(attr(kid(t, "period") ?? t, "val") ?? 2) } : {}),
        ...(flag(t, "dispEq", false) ? { equation: true } : {}), ...(flag(t, "dispRSqr", false) ? { r2: true } : {}) });
    }
  });
  if (trendlines.length) options.trendlines = trendlines;
  if (kind === "combo") {
    const second = used.map(({ g }, j) => (isSecondary(g) ? j : -1)).filter((j) => j >= 0);
    if (second.length) options.secondary = second;
  }
  return { kind, range: rectName(block.r), title, at: keyOf(at), ...(Object.keys(options).length ? { options: options as ChartOptions } : {}) };
}

function readEx(part: El, host: string, at: Addr): ChartOut | string {
  const chart = kid(part, "chart"), data = kid(part, "chartData");
  const series = chart ? all(chart, "series") : [];
  if (!chart || !data || !series.length) return "it has no series";
  const id = attr(series[0], "layoutId") ?? "";
  const kind = ({ waterfall: "waterfall", funnel: "funnel", treemap: "treemap", sunburst: "sunburst", boxWhisker: "box" } as Record<string, ChartKind>)[id]
    ?? (id === "clusteredColumn" && all(series[0], "binning").length ? "histogram" : undefined);
  if (!kind) return `a ${id === "regionMap" ? "map" : id || "chart"} of a kind the app does not have`;
  const used = kind === "box" ? series : series.slice(0, 1);
  const refs: string[] = [];
  for (const s of used) {
    const tx = all(s, "tx")[0], f = tx && all(tx, "f")[0];
    if (f) refs.push(textOf(f).trim());
    const dataId = attr(all(s, "dataId")[0] ?? s, "val");
    const block = all(data, "data").find((d) => attr(d, "id") === dataId);
    if (!block) return "a series without its data";
    for (const dim of [...kids(block, "strDim"), ...kids(block, "numDim")]) {
      const fx = kid(dim, "f");
      if (fx) refs.push(textOf(fx).trim());
      else if (attr(dim, "type") === "val") return "it holds its data itself, not cells";
    }
  }
  const block = blockOf(refs, host);
  if (typeof block === "string") return block;
  const options: Record<string, unknown> = {};
  const legend = kid(chart, "legend");
  if (!legend) options.legend = "none";
  else {
    const pos = attr(legend, "pos");
    if (pos === "t") options.legend = "top"; else if (pos === "r") options.legend = "right"; else if (pos === "l") options.legend = "left";
  }
  if (all(chart, "dataLabels").length) options.labels = true;
  const axes = all(chart, "axis");
  const vAxis = axes.find((a) => kid(a, "valScaling"));
  if (vAxis && !kid(vAxis, "majorGridlines")) options.gridlines = false;
  const catAxis = axes.find((a) => kid(a, "catScaling"));
  if (catAxis && titleOf(kid(catAxis, "title"))) options.xTitle = titleOf(kid(catAxis, "title"));
  if (vAxis && titleOf(kid(vAxis, "title"))) options.yTitle = titleOf(kid(vAxis, "title"));
  const firstVal = parseRef(refs.find((r) => !!parseRef(r)) ?? "");
  const f0 = all(data, "numDim")[0] && all(all(data, "numDim")[0], "f")[0];
  const vref = f0 ? parseRef(textOf(f0).trim()) : firstVal;
  if (vref && vref.r.a.row === vref.r.b.row && vref.r.a.col !== vref.r.b.col) options.byRows = true;
  return { kind, range: rectName(block.r), title: titleOf(kid(chart, "title")), at: keyOf(at), ...(Object.keys(options).length ? { options: options as ChartOptions } : {}) };
}

const rectName = (r: Rect): string => (r.a.col === r.b.col && r.a.row === r.b.row ? keyOf(r.a) : `${keyOf(r.a)}:${keyOf(r.b)}`);

// A drawing part's charts: each graphic frame holding one (inside an
// alternate content too, as the extended ones are), its part read by the
// drawing's relationships. `seen` collects the chart parts looked at; what
// cannot be read is said in `report`.
export function readCharts(part: El, rels: ReadonlyMap<string, string>, parts: ReadonlyMap<string, Uint8Array>, parse: (bytes: Uint8Array) => El, sheet: string, report: string[], seen: Set<string>): ChartOut[] {
  const out: ChartOut[] = [];
  for (const anchor of kids(part).filter((k) => /Anchor$/.test(local(k.name)))) {
    const frame = all(anchor, "graphicFrame")[0], ref = frame && all(frame, "chart")[0];
    if (!frame || !ref) continue;
    const from = kid(anchor, "from"), num = (n: string) => Number(textOf(from ? kid(from, n) ?? { name: "", attrs: {}, children: ["0"] } : { name: "", attrs: {}, children: ["0"] })) | 0;
    const nv = all(frame, "cNvPr")[0], name = (nv && attr(nv, "name")) || "a chart";
    const target = rels.get(attr(ref, "id") ?? ""), bytes = target ? parts.get(target) : undefined;
    if (!target || !bytes) { report.push(`${sheet}: the chart ${name} (its part not in the file) left out`); continue; }
    seen.add(target);
    let read: ChartOut | string;
    try { read = readChart(parse(bytes), sheet, { col: Math.max(0, num("col")), row: Math.max(0, num("row")) }); } catch (e) { read = "it could not be read"; }
    if (typeof read === "string") {
      const title = (() => { try { return titleOf(all(parse(bytes), "title")[0]); } catch (e) { return ""; } })();
      report.push(`${sheet}: the chart ${title || name} left out: ${read}`);
    } else out.push(read);
  }
  return out;
}
