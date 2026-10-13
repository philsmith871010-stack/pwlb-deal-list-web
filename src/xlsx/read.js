/* READING AN .XLSX INTO A PLAIN MODEL.
 *
 * Everything the workbook holds that a person would see or use: the sheets and
 * their order and visibility, every cell's value, formula, cached result and
 * style, column widths and row heights, merged cells, frozen panes, gridlines,
 * tab colours, defined names, data validation (dropdowns and their prompts),
 * conditional formatting, hyperlinks, Excel tables, sheet protection and the
 * charts with where they sit. The model knows nothing of any spreadsheet
 * library: src/univer/ turns it into one. The part paths and the original bytes
 * are kept so that saving can write the user's own workbook back with only the
 * typed-in cells changed (src/xlsx/write.js).
 */
import { unzip } from "./zip.js";
import { parseCell, parseRange, shiftFormula, cellName } from "./ref.js";

const dec = new TextDecoder();
const parser = () => new DOMParser();

function xml(parts, path){
  const b = parts.get(path);
  if(!b) return null;
  const doc = parser().parseFromString(dec.decode(b), "application/xml");
  if(doc.getElementsByTagName("parsererror").length) throw new Error("Could not read " + path);
  return doc;
}
/* Children and descendants by local name, whatever prefix the writer used. */
const kids = (el, name) => el ? [...el.children].filter(k => k.localName === name) : [];
const kid = (el, name) => el ? [...el.children].find(k => k.localName === name) || null : null;
const all = (el, name) => el ? [...el.getElementsByTagNameNS("*", name)] : [];
const attr = (el, name) => el ? el.getAttribute(name) : null;
const num = (el, name, d = null) => { const v = attr(el, name); return v == null || v === "" ? d : +v; };
const bool = (el, name, d = false) => { const v = attr(el, name); return v == null ? d : v === "1" || v === "true"; };
const rattr = el => el ? (el.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || el.getAttribute("r:id")) : null;

function resolve(base, target){
  if(target.startsWith("/")) return target.slice(1);
  const dir = base.split("/").slice(0, -1);
  for(const p of target.split("/")){ if(p === "..") dir.pop(); else if(p !== ".") dir.push(p); }
  return dir.join("/");
}
function rels(parts, path){
  const relPath = path.replace(/([^/]+)$/, "_rels/$1.rels");
  const doc = xml(parts, relPath);
  const out = {};
  if(!doc) return out;
  for(const r of all(doc, "Relationship")){
    const external = attr(r, "TargetMode") === "External";
    const t = attr(r, "Target");
    out[attr(r, "Id")] = {type: attr(r, "Type"), target: t, path: external ? null : resolve(path, t), external};
  }
  return out;
}

/* ---- colours: rgb, theme with tint, indexed ---- */
const INDEXED = ["000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF","000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF",
  "800000","008000","000080","808000","800080","008080","C0C0C0","808080","9999FF","993366","FFFFCC","CCFFFF","660066","FF8080","0066CC","CCCCFF",
  "000080","FF00FF","FFFF00","00FFFF","800080","800000","008080","0000FF","00CCFF","CCFFFF","CCFFCC","FFFF99","99CCFF","FF99CC","CC99FF","FFCC99",
  "3366FF","33CCCC","99CC00","FFCC00","FF9900","FF6600","666699","969696","003366","339966","003300","333300","993300","993366","333399","333333"];
function rgbToHsl(r, g, b){
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
  if(mx === mn) return [0, 0, l];
  const d = mx - mn, s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
  const h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}
function hslToRgb(h, s, l){
  if(!s) return [l, l, l].map(v => Math.round(v * 255));
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const f = t => { t = (t + 1) % 1; return t < 1/6 ? p + (q - p) * 6 * t : t < 1/2 ? q : t < 2/3 ? p + (q - p) * (2/3 - t) * 6 : p; };
  return [f(h + 1/3), f(h), f(h - 1/3)].map(v => Math.round(v * 255));
}
function applyTint(hex, tint){
  if(!tint) return hex;
  const [h, s, l] = rgbToHsl(parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16));
  const l2 = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return hslToRgb(h, s, l2).map(v => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}
function colour(el, theme){
  if(!el) return null;
  if(attr(el, "auto") === "1") return null;
  let hex = null;
  const rgb = attr(el, "rgb");
  if(rgb) hex = rgb.length === 8 ? rgb.slice(2) : rgb;
  else if(attr(el, "theme") != null) hex = theme[+attr(el, "theme")] || null;
  else if(attr(el, "indexed") != null){ const i = +attr(el, "indexed"); hex = i === 64 ? null : INDEXED[i] || null; }
  if(!hex) return null;
  return "#" + applyTint(hex.toUpperCase(), num(el, "tint", 0));
}

function readTheme(parts, wbPath, wbRels){
  const rel = Object.values(wbRels).find(r => /\/theme$/.test(r.type));
  const doc = rel ? xml(parts, rel.path) : null;
  const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const def = {lt1: "FFFFFF", dk1: "000000", lt2: "E7E6E6", dk2: "44546A", accent1: "4472C4", accent2: "ED7D31", accent3: "A5A5A5", accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47", hlink: "0563C1", folHlink: "954F72"};
  const scheme = doc ? all(doc, "clrScheme")[0] : null;
  const got = {};
  if(scheme) for(const c of scheme.children){
    const v = kid(c, "srgbClr") || kid(c, "sysClr");
    if(v) got[c.localName] = (attr(v, "val") && v.localName === "srgbClr") ? attr(v, "val") : attr(v, "lastClr");
  }
  return order.map(k => (got[k] || def[k]).toUpperCase());
}

/* ---- styles ---- */
const BUILTIN_FMT = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
  14: "dd/mm/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "dd/mm/yyyy hh:mm",
  37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@"};

function readFont(f, theme){
  if(!f) return {};
  const v = (name) => kid(f, name);
  const has = (name) => { const e = v(name); return e ? attr(e, "val") !== "0" && attr(e, "val") !== "false" : false; };
  const u = v("u");
  return {
    name: attr(v("name"), "val") || null, size: num(v("sz"), "val"), bold: has("b"), italic: has("i"), strike: has("strike"),
    underline: u ? (attr(u, "val") || "single") : null, color: colour(v("color"), theme), vertAlign: attr(v("vertAlign"), "val")
  };
}
function readFill(f, theme){
  if(!f) return null;
  const pf = kid(f, "patternFill");
  if(pf){
    const type = attr(pf, "patternType");
    if(!type || type === "none") return null;
    const fg = colour(kid(pf, "fgColor"), theme), bg = colour(kid(pf, "bgColor"), theme);
    return type === "solid" ? (fg || bg) : (fg || bg);
  }
  const gf = kid(f, "gradientFill");
  if(gf){ const st = all(gf, "stop")[0]; return st ? colour(kid(st, "color"), theme) : null; }
  return null;
}
function readBorder(b, theme){
  if(!b) return null;
  const out = {};
  for(const side of ["left", "right", "top", "bottom", "diagonal"]){
    const e = kid(b, side) || (side === "left" ? kid(b, "start") : side === "right" ? kid(b, "end") : null);
    const st = attr(e, "style");
    if(e && st && st !== "none") out[side] = {style: st, color: colour(kid(e, "color"), theme) || "#000000"};
  }
  if(out.diagonal){ out.diagUp = bool(b, "diagonalUp"); out.diagDown = bool(b, "diagonalDown"); }
  return Object.keys(out).length ? out : null;
}
function readXf(x, fmts, fonts, fills, borders){
  const al = kid(x, "alignment"), pr = kid(x, "protection");
  const fmtId = num(x, "numFmtId", 0);
  return {
    numFmt: fmts[fmtId] != null ? fmts[fmtId] : (BUILTIN_FMT[fmtId] || "General"),
    font: fonts[num(x, "fontId", 0)] || {}, fill: fills[num(x, "fillId", 0)] || null, border: borders[num(x, "borderId", 0)] || null,
    align: al ? {h: attr(al, "horizontal"), v: attr(al, "vertical"), wrap: bool(al, "wrapText"), indent: num(al, "indent", 0),
      rotate: num(al, "textRotation", 0), shrink: bool(al, "shrinkToFit")} : null,
    locked: pr ? attr(pr, "locked") !== "0" : true, hidden: pr ? bool(pr, "hidden") : false
  };
}
function readStyles(parts, path, theme){
  const doc = xml(parts, path);
  const fmts = {};
  if(!doc) return {xfs: [{numFmt: "General", font: {}, fill: null, border: null, align: null, locked: true}], dxfs: [], fmts};
  for(const f of all(doc, "numFmt")) fmts[+attr(f, "numFmtId")] = attr(f, "formatCode");
  const fonts = kids(kid(doc.documentElement, "fonts"), "font").map(f => readFont(f, theme));
  const fills = kids(kid(doc.documentElement, "fills"), "fill").map(f => readFill(f, theme));
  const borders = kids(kid(doc.documentElement, "borders"), "border").map(b => readBorder(b, theme));
  const xfs = kids(kid(doc.documentElement, "cellXfs"), "xf").map(x => readXf(x, fmts, fonts, fills, borders));
  const dxfs = kids(kid(doc.documentElement, "dxfs"), "dxf").map(d => {
    const nf = kid(d, "numFmt");
    const fill = kid(d, "fill"), pf = kid(fill, "patternFill");
    /* In a differential fill a solid colour is written as bgColor by Excel and as fgColor by others. */
    const bg = pf ? (colour(kid(pf, "bgColor"), theme) || colour(kid(pf, "fgColor"), theme)) : null;
    return {font: kid(d, "font") ? readFont(kid(d, "font"), theme) : null, fill: bg, border: readBorder(kid(d, "border"), theme), numFmt: nf ? attr(nf, "formatCode") : null};
  });
  return {xfs, dxfs, fmts, fonts};
}

/* ---- text ---- */
function richText(si){
  if(!si) return "";
  const t = kid(si, "t");
  if(t && !kid(si, "r")) return t.textContent;
  return kids(si, "r").map(r => (kid(r, "t") || {}).textContent || "").join("") || (t ? t.textContent : "");
}

/* ---- the sheets ---- */
function readSheet(parts, path, info, ctx){
  const doc = xml(parts, path);
  const root = doc.documentElement;
  const sheet = {name: info.name, state: info.state, path, cells: new Map(), rows: new Map(), cols: [], merges: [],
    cf: [], dv: [], links: [], tables: [], charts: [], protection: null, maxRow: 0, maxCol: 0};
  const pr = kid(root, "sheetPr");
  sheet.tabColor = colour(kid(pr, "tabColor"), ctx.theme);
  const view = kid(kid(root, "sheetViews"), "sheetView");
  sheet.showGridLines = view ? attr(view, "showGridLines") !== "0" : true;
  sheet.zoom = num(view, "zoomScale", 100);
  sheet.rightToLeft = bool(view, "rightToLeft");
  const pane = kid(view, "pane");
  sheet.frozen = pane && /frozen/.test(attr(pane, "state") || "") ? {cols: num(pane, "xSplit", 0), rows: num(pane, "ySplit", 0)} : null;
  const fmt = kid(root, "sheetFormatPr");
  sheet.defaultRowHeight = num(fmt, "defaultRowHeight", 15);
  sheet.defaultColWidth = num(fmt, "defaultColWidth", null);
  sheet.baseColWidth = num(fmt, "baseColWidth", 8);
  for(const c of kids(kid(root, "cols"), "col"))
    sheet.cols.push({min: num(c, "min") - 1, max: num(c, "max") - 1, width: num(c, "width"), hidden: bool(c, "hidden"), style: num(c, "style", null), custom: bool(c, "customWidth")});

  /* the cells, and Excel's shared formulas spelled out cell by cell */
  const shared = {};
  for(const row of kids(kid(root, "sheetData"), "row")){
    const r = num(row, "r") - 1;
    const rd = {};
    if(attr(row, "ht") != null) rd.ht = num(row, "ht");
    if(bool(row, "hidden")) rd.hidden = true;
    if(bool(row, "customFormat") && attr(row, "s") != null) rd.style = num(row, "s");
    if(Object.keys(rd).length) sheet.rows.set(r, rd);
    let cAuto = 0;
    for(const c of kids(row, "c")){
      const ref = attr(c, "r");
      const at = ref ? parseCell(ref) : {r, c: cAuto};
      cAuto = at.c + 1;
      const t = attr(c, "t") || "n", s = num(c, "s", 0);
      const ve = kid(c, "v"), fe = kid(c, "f");
      const cell = {s};
      let v = ve ? ve.textContent : null;
      if(t === "s") v = v == null ? null : ctx.sst[+v];
      else if(t === "inlineStr") v = richText(kid(c, "is"));
      else if(t === "b") v = v == null ? null : v === "1";
      else if(t === "e") cell.error = v;
      else if(t === "str") v = v == null ? "" : v;
      else if(t === "d") v = v == null ? null : isoToSerial(v, ctx.date1904);
      else if(v != null && v !== "") v = +v;
      else v = null;
      if(t === "e"){ cell.v = v; cell.type = "error"; }
      else if(v !== null){ cell.v = v; cell.type = typeof v === "number" ? "n" : typeof v === "boolean" ? "b" : "s"; }
      if(fe){
        const ft = attr(fe, "t");
        let f = fe.textContent;
        if(ft === "shared"){
          const si = attr(fe, "si");
          if(f) shared[si] = {f, r: at.r, c: at.c};
          else if(shared[si]) f = shiftFormula(shared[si].f, at.r - shared[si].r, at.c - shared[si].c);
        }
        if(ft === "array") cell.arrayRef = attr(fe, "ref");
        if(f) cell.f = f.replace(/_xlfn\.|_xlws\./g, "");
        if(cell.f != null) cell.cached = cell.v === undefined ? null : cell.v;
      }
      if(cell.v !== undefined || cell.f != null || s){
        let rowMap = sheet.cells.get(at.r);
        if(!rowMap){ rowMap = new Map(); sheet.cells.set(at.r, rowMap); }
        rowMap.set(at.c, cell);
        if(cell.v !== undefined || cell.f != null){ if(at.r > sheet.maxRow) sheet.maxRow = at.r; if(at.c > sheet.maxCol) sheet.maxCol = at.c; }
      }
    }
  }

  for(const m of all(kid(root, "mergeCells"), "mergeCell")){ const rg = parseRange(attr(m, "ref")); if(rg) sheet.merges.push(rg); }

  for(const cf of kids(root, "conditionalFormatting")){
    const ranges = (attr(cf, "sqref") || "").split(/\s+/).filter(Boolean).map(a => parseRange(a)).filter(Boolean);
    for(const rule of kids(cf, "cfRule")){
      sheet.cf.push({ranges, type: attr(rule, "type"), priority: num(rule, "priority", 0), dxf: num(rule, "dxfId", null),
        operator: attr(rule, "operator"), text: attr(rule, "text"), stopIfTrue: bool(rule, "stopIfTrue"),
        formulas: kids(rule, "formula").map(f => f.textContent)});
    }
  }

  for(const d of kids(kid(root, "dataValidations"), "dataValidation")){
    sheet.dv.push({ranges: (attr(d, "sqref") || "").split(/\s+/).filter(Boolean).map(a => parseRange(a)).filter(Boolean),
      type: attr(d, "type") || "any", operator: attr(d, "operator") || "between", allowBlank: bool(d, "allowBlank"),
      showDropDown: !bool(d, "showDropDown"),   /* Excel's flag is inverted: 1 hides the arrow */
      showInputMessage: bool(d, "showInputMessage"), showErrorMessage: bool(d, "showErrorMessage"), errorStyle: attr(d, "errorStyle") || "stop",
      errorTitle: attr(d, "errorTitle"), error: attr(d, "error"), promptTitle: attr(d, "promptTitle"), prompt: attr(d, "prompt"),
      formula1: (kid(d, "formula1") || {}).textContent ?? null, formula2: (kid(d, "formula2") || {}).textContent ?? null});
  }

  const srels = rels(parts, path);
  for(const h of kids(kid(root, "hyperlinks"), "hyperlink")){
    const rg = parseRange(attr(h, "ref"));
    const rel = srels[rattr(h)];
    let target = rel ? rel.target : null, location = attr(h, "location");
    if(target && target.startsWith("#")){ location = target.slice(1); target = null; }
    sheet.links.push({range: rg, target, location, tooltip: attr(h, "tooltip"), display: attr(h, "display")});
  }

  const sp = kid(root, "sheetProtection");
  if(sp && bool(sp, "sheet")){
    sheet.protection = {};
    for(const a of sp.attributes) sheet.protection[a.name] = a.value === "1" || a.value === "true" ? true : a.value === "0" || a.value === "false" ? false : a.value;
  }

  for(const tp of kids(kid(root, "tableParts"), "tablePart")){
    const rel = srels[rattr(tp)];
    const tdoc = rel ? xml(parts, rel.path) : null;
    if(!tdoc) continue;
    const t = tdoc.documentElement, si = kid(t, "tableStyleInfo");
    sheet.tables.push({path: rel.path, name: attr(t, "displayName") || attr(t, "name"), ref: parseRange(attr(t, "ref")), refText: attr(t, "ref"),
      headerRows: num(t, "headerRowCount", 1), totalsRows: num(t, "totalsRowCount", 0), autoFilter: !!kid(t, "autoFilter"),
      columns: all(t, "tableColumn").map(c => attr(c, "name")),
      style: si ? {name: attr(si, "name"), rowStripes: bool(si, "showRowStripes"), colStripes: bool(si, "showColumnStripes"),
        firstCol: bool(si, "showFirstColumn"), lastCol: bool(si, "showLastColumn")} : null});
  }

  const dr = kid(root, "drawing");
  if(dr && srels[rattr(dr)]) sheet.charts = readDrawing(parts, srels[rattr(dr)].path, ctx);
  return sheet;
}

function isoToSerial(iso, d1904){
  const t = Date.parse(iso.length <= 10 ? iso + "T00:00:00Z" : iso + (iso.endsWith("Z") ? "" : "Z"));
  const base = d1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (t - base) / 86400000;
}

/* ---- drawings and charts ---- */
const EMU_PER_PX = 9525;
function readDrawing(parts, path, ctx){
  const doc = xml(parts, path);
  const drels = rels(parts, path);
  const out = [];
  for(const a of [...all(doc, "twoCellAnchor"), ...all(doc, "oneCellAnchor"), ...all(doc, "absoluteAnchor")]){
    const chartEl = all(a, "chart")[0];
    if(!chartEl) continue;
    const rel = drels[rattr(chartEl)];
    if(!rel) continue;
    const pos = e => e ? {c: +kid(e, "col").textContent, cOff: +kid(e, "colOff").textContent / EMU_PER_PX,
      r: +kid(e, "row").textContent, rOff: +kid(e, "rowOff").textContent / EMU_PER_PX} : null;
    const ext = kid(a, "ext");
    const anchor = {type: a.localName, from: pos(kid(a, "from")), to: pos(kid(a, "to")),
      ext: ext ? {w: num(ext, "cx") / EMU_PER_PX, h: num(ext, "cy") / EMU_PER_PX} : null,
      abs: kid(a, "pos") ? {x: num(kid(a, "pos"), "x") / EMU_PER_PX, y: num(kid(a, "pos"), "y") / EMU_PER_PX} : null};
    const nv = all(a, "cNvPr")[0];
    out.push({anchor, name: attr(nv, "name"), path: rel.path, chart: readChart(parts, rel.path, ctx)});
  }
  return out;
}

function shapeColour(spPr, theme){
  if(!spPr) return null;
  const sf = kid(spPr, "solidFill");
  return sf ? drawingColour(sf, theme) : kid(spPr, "noFill") ? "none" : null;
}
function drawingColour(holder, theme){
  if(!holder) return null;
  const s = kid(holder, "srgbClr");
  if(s) return "#" + attr(s, "val").toUpperCase();
  const sc = kid(holder, "schemeClr");
  if(sc){
    const map = {bg1: 0, tx1: 1, bg2: 2, tx2: 3, lt1: 0, dk1: 1, lt2: 2, dk2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11};
    const i = map[attr(sc, "val")];
    return i == null ? null : "#" + theme[i];
  }
  return null;
}
function lineProps(spPr, theme){
  const ln = kid(spPr, "ln");
  if(!ln) return null;
  if(kid(ln, "noFill")) return {none: true};
  return {width: num(ln, "w", 9525) / 12700, color: drawingColour(kid(ln, "solidFill"), theme), dash: attr(kid(ln, "prstDash"), "val") || "solid"};
}
function textRuns(tx, theme){
  const rich = kid(tx, "rich");
  if(!rich) return [];
  return kids(rich, "p").map(p => {
    const runs = kids(p, "r");
    const def = kid(kid(p, "pPr"), "defRPr");
    const rp = runs.length ? kid(runs[0], "rPr") : null;
    const props = rp || def;
    return {text: runs.map(r => (kid(r, "t") || {}).textContent || "").join(""), size: (num(props, "sz", 1000)) / 100,
      bold: bool(props, "b"), color: drawingColour(kid(props, "solidFill"), theme), font: attr(kid(props, "latin"), "typeface"),
      align: attr(kid(p, "pPr"), "algn")};
  }).filter(p => p.text);
}
function seriesRef(e){
  if(!e) return null;
  const ref = kid(e, "numRef") || kid(e, "strRef");
  const f = ref ? (kid(ref, "f") || {}).textContent : null;
  const cacheEl = ref ? (kid(ref, "numCache") || kid(ref, "strCache")) : (kid(e, "numLit") || kid(e, "strLit"));
  const cache = cacheEl ? kids(cacheEl, "pt").map(pt => ({i: num(pt, "idx"), v: (kid(pt, "v") || {}).textContent})) : [];
  const fmt = cacheEl ? (kid(cacheEl, "formatCode") || {}).textContent : null;
  return {f, cache, fmt};
}
function readChart(parts, path, ctx){
  const doc = xml(parts, path);
  const theme = ctx.theme;
  const cs = doc.documentElement, chart = kid(cs, "chart");
  const out = {title: [], plots: [], axes: {}, legend: null, plotLayout: null};
  const title = kid(chart, "title");
  if(title){
    out.title = textRuns(kid(title, "tx"), theme);
    const ml = kid(kid(title, "layout"), "manualLayout");
    if(ml) out.titleLayout = {x: num(kid(ml, "x"), "val", 0), y: num(kid(ml, "y"), "val", 0)};
  }
  out.autoTitleDeleted = bool(kid(chart, "autoTitleDeleted"), "val");
  const pa = kid(chart, "plotArea");
  const pml = kid(kid(pa, "layout"), "manualLayout");
  if(pml) out.plotLayout = {x: num(kid(pml, "x"), "val"), y: num(kid(pml, "y"), "val"), w: num(kid(pml, "w"), "val"), h: num(kid(pml, "h"), "val"),
    target: attr(kid(pml, "layoutTarget"), "val") || "outer"};
  for(const p of pa.children){
    const m = /^(bar|line|pie|doughnut|area|scatter|radar|bar3D|line3D|pie3D|area3D)Chart$/.exec(p.localName);
    if(!m) continue;
    const plot = {type: m[1].replace("3D", ""), barDir: attr(kid(p, "barDir"), "val") || "col", grouping: attr(kid(p, "grouping"), "val") || "clustered",
      varyColors: bool(kid(p, "varyColors"), "val"), gapWidth: num(kid(p, "gapWidth"), "val", 150), overlap: num(kid(p, "overlap"), "val", 0),
      holeSize: num(kid(p, "holeSize"), "val", null), firstSliceAng: num(kid(p, "firstSliceAng"), "val", 0),
      axIds: kids(p, "axId").map(a => attr(a, "val")), series: [], dLbls: readDLbls(kid(p, "dLbls"))};
    for(const s of kids(p, "ser")){
      const tx = kid(s, "tx");
      const sp = kid(s, "spPr");
      const ser = {idx: num(kid(s, "idx"), "val"), order: num(kid(s, "order"), "val"),
        name: tx ? (kid(tx, "v") ? kid(tx, "v").textContent : seriesRef(tx)) : null,
        cat: seriesRef(kid(s, "cat") || kid(s, "xVal")), val: seriesRef(kid(s, "val") || kid(s, "yVal")),
        color: shapeColour(sp, theme), line: lineProps(sp, theme), smooth: bool(kid(s, "smooth"), "val"),
        marker: (() => { const mk = kid(s, "marker"); if(!mk) return null; return {symbol: attr(kid(mk, "symbol"), "val") || "auto", size: num(kid(mk, "size"), "val", 5), color: shapeColour(kid(mk, "spPr"), theme)}; })(),
        points: {}, dLbls: readDLbls(kid(s, "dLbls"))};
      for(const dp of kids(s, "dPt")) ser.points[num(kid(dp, "idx"), "val")] = shapeColour(kid(dp, "spPr"), theme);
      plot.series.push(ser);
    }
    out.plots.push(plot);
  }
  for(const ax of pa.children){
    if(!/^(catAx|valAx|dateAx|serAx)$/.test(ax.localName)) continue;
    const sc = kid(ax, "scaling");
    const nf = kid(ax, "numFmt");
    out.axes[attr(kid(ax, "axId"), "val")] = {kind: ax.localName, deleted: bool(kid(ax, "delete"), "val"), pos: attr(kid(ax, "axPos"), "val"),
      reverse: attr(kid(sc, "orientation"), "val") === "maxMin", min: num(kid(sc, "min"), "val"), max: num(kid(sc, "max"), "val"),
      majorUnit: num(kid(ax, "majorUnit"), "val"), numFmt: nf ? attr(nf, "formatCode") : null, gridlines: !!kid(ax, "majorGridlines"),
      gridColor: lineProps(kid(kid(ax, "majorGridlines"), "spPr"), theme), title: kid(ax, "title") ? textRuns(kid(kid(ax, "title"), "tx"), theme) : [],
      tickLblPos: attr(kid(ax, "tickLblPos"), "val") || "nextTo", crossAx: attr(kid(ax, "crossAx"), "val"),
      labelFont: (() => { const tp = kid(ax, "txPr"); const d = tp ? all(tp, "defRPr")[0] : null; return d ? {size: num(d, "sz", 1000) / 100, color: drawingColour(kid(d, "solidFill"), theme)} : null; })(),
      line: lineProps(kid(ax, "spPr"), theme), tickLblSkip: num(kid(ax, "tickLblSkip"), "val")};
  }
  const lg = kid(chart, "legend");
  if(lg) out.legend = {pos: attr(kid(lg, "legendPos"), "val") || "r", overlay: bool(kid(lg, "overlay"), "val"),
    font: (() => { const d = all(lg, "defRPr")[0]; return d ? {size: num(d, "sz", 1000) / 100, color: drawingColour(kid(d, "solidFill"), theme)} : null; })(),
    hidden: kids(lg, "legendEntry").filter(e => bool(kid(e, "delete"), "val")).map(e => num(kid(e, "idx"), "val"))};
  return out;
}
function readDLbls(d){
  if(!d) return null;
  if(bool(kid(d, "delete"), "val")) return {show: false};
  const nf = kid(d, "numFmt");
  return {show: bool(kid(d, "showVal"), "val"), pos: attr(kid(d, "dLblPos"), "val"), numFmt: nf ? attr(nf, "formatCode") : null,
    showPercent: bool(kid(d, "showPercent"), "val"), showCat: bool(kid(d, "showCatName"), "val")};
}

/* ---- the workbook ---- */
/* opts.only: the names of the sheets to read (the rest are listed but not parsed). */
export async function readXlsx(bytes, opts = {}){
  const parts = await unzip(bytes);
  const ct = xml(parts, "[Content_Types].xml");
  const rootRels = rels(parts, "");
  const wbRel = Object.values(rootRels).find(r => /officeDocument$/.test(r.type));
  const wbPath = wbRel ? wbRel.path : "xl/workbook.xml";
  const wb = xml(parts, wbPath);
  if(!wb) throw new Error("This workbook has no workbook part.");
  const wbRels = rels(parts, wbPath);
  const theme = readTheme(parts, wbPath, wbRels);
  const date1904 = bool(kid(wb.documentElement, "workbookPr"), "date1904");
  const stylesRel = Object.values(wbRels).find(r => /\/styles$/.test(r.type));
  const styles = readStyles(parts, stylesRel ? stylesRel.path : "xl/styles.xml", theme);
  const sstRel = Object.values(wbRels).find(r => /\/sharedStrings$/.test(r.type));
  const sstDoc = sstRel ? xml(parts, sstRel.path) : null;
  const sst = sstDoc ? kids(sstDoc.documentElement, "si").map(richText) : [];
  const ctx = {theme, styles, sst, date1904};
  const infos = all(wb, "sheet").map(s => ({name: attr(s, "name"), sheetId: attr(s, "sheetId"), state: attr(s, "state") || "visible", rel: wbRels[rattr(s)]}));
  const want = opts.only ? new Set(opts.only.map(n => n.toLowerCase())) : null;
  const allSheets = infos.filter(i => i.rel && /\/worksheet$/.test(i.rel.type));
  const sheets = allSheets.filter(i => !want || want.has(i.name.toLowerCase())).map(i => readSheet(parts, i.rel.path, i, ctx));
  const names = all(wb, "definedName").map(d => ({name: attr(d, "name"), local: attr(d, "localSheetId") == null ? null : +attr(d, "localSheetId"),
    ref: d.textContent, hidden: bool(d, "hidden")}));
  const view = all(wb, "workbookView")[0];
  const active = allSheets[num(view, "activeTab", 0)];
  return {bytes: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), parts, wbPath, theme, date1904, styles, sheets, names,
    allSheetNames: allSheets.map(i => i.name), activeName: active ? active.name : null, contentTypes: ct,
    lockStructure: bool(kid(wb.documentElement, "workbookProtection"), "lockStructure")};
}

export { cellName };
