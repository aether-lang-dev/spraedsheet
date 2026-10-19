// A sheet's drawings in an XLSX file (DrawingML's spreadsheet drawing
// part): pictures (xdr:pic, their bytes a media part by relationship),
// shapes and text boxes (xdr:sp, a preset geometry, a fill and a line,
// their text), each a one-cell anchor at its cell, of a size in EMUs
// (9,525 to a pixel). Written from the app's drawings; read back, a
// two-cell or absolute anchor taken at its shape's own size.
import { attr, all, escapeXml, kid, kids, textOf } from "./xml.ts";
import type { El } from "./xml.ts";
import { COLS, ROWS } from "./address.ts";
import { EXTENSION, PRESET, SHAPES } from "./drawings.ts";
import type { Drawing, ShapeKind } from "./drawings.ts";
import { fromBase64, toBase64 } from "./csv.ts";

const EMU = 9525;
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
// A theme's roles as DrawingML's scheme colours, and back.
const SCHEME: Readonly<Record<string, string>> = { dark1: "dk1", light1: "lt1", dark2: "dk2", light2: "lt2", accent1: "accent1", accent2: "accent2", accent3: "accent3", accent4: "accent4", accent5: "accent5", accent6: "accent6", hyperlink: "hlink" };
const ROLE_OF: Readonly<Record<string, string>> = { ...Object.fromEntries(Object.entries(SCHEME).map(([k, v]) => [v, k])), tx1: "dark1", bg1: "light1", tx2: "dark2", bg2: "light2" };

const colourXml = (c: string) => (SCHEME[c] ? `<a:schemeClr val="${SCHEME[c]}"/>` : `<a:srgbClr val="${c.replace("#", "").toUpperCase()}"/>`);

// The drawing part for some drawings, and its pictures: each picture's
// media (its bytes, its part's name) and the relationship naming it.
// A chart in the drawing: its relationship, its name, its cell, and whether
// its part is an extended one (Office 2016's chartex).
export interface ChartFrame { readonly rId: string; readonly name: string; readonly at: Addr; readonly ex: boolean }
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CHARTEX_URI = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const CHART_W = 400, CHART_H = 240;

function chartAnchor(f: ChartFrame, id: number): string {
  const cx = CHART_W * EMU, cy = CHART_H * EMU;
  const frame = (inner: string) => `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeXml(f.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>${inner}</a:graphic></xdr:graphicFrame>`;
  const body = f.ex
    ? `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">${
      frame(`<a:graphicData uri="${CHARTEX_URI}"><cx:chart xmlns:cx="${CHARTEX_URI}" xmlns:r="${R}" r:id="${f.rId}"/></a:graphicData>`)}</mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${escapeXml(f.name)}"/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>This chart isn't available in your version of Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>`
    : frame(`<a:graphicData uri="${CHART_URI}"><c:chart xmlns:c="${CHART_URI}" xmlns:r="${R}" r:id="${f.rId}"/></a:graphicData>`);
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${f.at.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${f.at.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/>${body}<xdr:clientData/></xdr:oneCellAnchor>`;
}

export function drawingXml(ds: readonly Drawing[], nextMedia: () => number, charts: readonly ChartFrame[] = []): { xml: string; media: { rId: string; name: string; bytes: Uint8Array }[] } {
  const media: { rId: string; name: string; bytes: Uint8Array }[] = [];
  const anchors = ds.map((d, i) => {
    const id = i + 2, cx = d.w * EMU, cy = d.h * EMU;
    const xfrm = `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
    let body: string;
    if (d.kind === "picture") {
      const rId = `rId${media.length + 1}`, ext = EXTENSION[d.mime ?? "image/png"] ?? "png";
      media.push({ rId, name: `xl/media/image${nextMedia()}.${ext}`, bytes: fromBase64(d.picture ?? "") });
      body = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${escapeXml(d.name)}"${d.alt ? ` descr="${escapeXml(d.alt)}"` : ""}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
    } else {
      const box = d.kind === "textbox", prst = box ? "rect" : PRESET[d.shape ?? "rectangle"];
      const fill = d.fill ?? (box ? "light1" : "accent1"), line = d.line ?? (box ? "#7f7f7f" : "#2f528f");
      const paras = (d.text ?? "").split("\n").map((l) => (l ? `<a:p><a:r><a:t>${escapeXml(l)}</a:t></a:r></a:p>` : "<a:p/>")).join("");
      body = `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${escapeXml(d.name)}"${d.alt ? ` descr="${escapeXml(d.alt)}"` : ""}/><xdr:cNvSpPr${box ? ' txBox="1"' : ""}/></xdr:nvSpPr><xdr:spPr>${xfrm}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${
        d.shape === "line" ? "<a:noFill/>" : `<a:solidFill>${colourXml(fill)}</a:solidFill>`}<a:ln w="${d.shape === "line" ? 19050 : 9525}"><a:solidFill>${colourXml(line)}</a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="${box ? "t" : "ctr"}"/><a:lstStyle/>${paras || "<a:p/>"}</xdr:txBody></xdr:sp>`;
    }
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${d.at.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${d.at.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/>${body}<xdr:clientData/></xdr:oneCellAnchor>`;
  });
  const frames = charts.map((f, i) => chartAnchor(f, ds.length + i + 2));
  return { xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}">${anchors.join("")}${frames.join("")}</xdr:wsDr>`, media };
}

// A drawing part read: its pictures (their bytes by `media`, from the
// part's relationships), shapes and text boxes; what it held that the app
// does not draw, said in `report`. The media parts it used go in `used`.
export function readDrawings(part: El, rels: ReadonlyMap<string, string>, media: (path: string) => Uint8Array | undefined, sheet: string, report: string[], used: Set<string>): Drawing[] {
  const out: Drawing[] = [];
  const names = new Set<string>();
  for (const anchor of kids(part).filter((k) => /Anchor$/.test(k.name))) {
    const from = kid(anchor, "from"), ext = kid(anchor, "ext");
    const col = Math.min(COLS - 1, Number(textOf(kid(from ?? anchor, "col") ?? { name: "", attrs: {}, children: ["0"] })) | 0);
    const row = Math.min(ROWS - 1, Number(textOf(kid(from ?? anchor, "row") ?? { name: "", attrs: {}, children: ["0"] })) | 0);
    const thing = kids(anchor).find((k) => ["pic", "sp", "cxnSp", "graphicFrame", "grpSp"].includes(k.name.slice(k.name.indexOf(":") + 1)));
    if (!thing) continue;
    const kind = thing.name.slice(thing.name.indexOf(":") + 1);
    if (kind === "graphicFrame") continue;   // a chart: said with the charts
    const nv = all(thing, "cNvPr")[0], name0 = (nv && attr(nv, "name")) || `Drawing ${out.length + 1}`, alt = nv ? attr(nv, "descr") : undefined;
    if (kind === "grpSp") { report.push(`${sheet}: the group ${name0} left out`); continue; }
    let name = name0;
    for (let i = 2; names.has(name.toLowerCase()); i++) name = `${name0} ${i}`;
    const sized = all(thing, "xfrm").map((x) => kid(x, "ext")).find((x) => x) ?? ext;
    const w = Math.max(1, Math.round(Number(sized ? attr(sized, "cx") : 952500) / EMU)), h = Math.max(1, Math.round(Number(sized ? attr(sized, "cy") : 952500) / EMU));
    const base = { name, at: { col, row }, w, h, ...(alt ? { alt } : {}) };
    if (kind === "pic") {
      const blip = all(thing, "blip")[0], target = blip ? rels.get(attr(blip, "embed") ?? "") : undefined, bytes = target ? media(target) : undefined;
      if (!target || !bytes) { report.push(`${sheet}: the picture ${name} (its bytes not in the file) left out`); continue; }
      const mime = Object.entries(EXTENSION).find(([, e]) => target.toLowerCase().endsWith(`.${e}`) || (e === "jpeg" && /\.jpg$/i.test(target)))?.[0];
      if (!mime) { report.push(`${sheet}: the picture ${name} (${target.split(".").pop()}) left out`); continue; }
      used.add(target);
      out.push({ kind: "picture", ...base, picture: toBase64(bytes), mime });
    } else {
      const geom = all(thing, "prstGeom")[0], prst = geom ? attr(geom, "prst") : "rect", box = all(thing, "cNvSpPr").some((x) => attr(x, "txBox") === "1");
      const shape = (Object.entries(PRESET).find(([, v]) => v === prst)?.[0] ?? (kind === "cxnSp" && /connector|line/i.test(prst ?? "") ? "line" : undefined)) as ShapeKind | undefined;
      if (!box && (!shape || !(SHAPES as readonly string[]).includes(shape))) { report.push(`${sheet}: the shape ${name} (${prst ?? "a custom geometry"}) left out`); continue; }
      const spPr = kid(thing, "spPr"), colourIn = (e: El | undefined): string | undefined => {
        const c = e && (kid(e, "srgbClr") ?? kid(e, "schemeClr"));
        if (!c) return undefined;
        const v = attr(c, "val") ?? "";
        return c.name.endsWith("srgbClr") ? `#${v.toLowerCase()}` : ROLE_OF[v];
      };
      const fill = colourIn(spPr && kid(spPr, "solidFill")), lineEl = spPr && kid(spPr, "ln"), line = colourIn(lineEl && kid(lineEl, "solidFill"));
      const body = kid(thing, "txBody"), text = body ? kids(body, "p").map((p) => all(p, "t").map(textOf).join("")).join("\n").replace(/\n+$/, "") : "";
      out.push({ kind: box ? "textbox" : "shape", ...base, ...(box ? {} : { shape }), ...(fill ? { fill } : {}), ...(line ? { line } : {}), ...(text ? { text } : {}) } as Drawing);
    }
    names.add(name.toLowerCase());
  }
  return out;
}
