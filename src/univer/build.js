/* THE MODEL, AS UNIVER'S WORKBOOK DATA.
 *
 * One IWorkbookData snapshot carries the sheets, cells, styles, sizes, merges,
 * frozen panes, names, data validation, conditional formatting and tables, so
 * the workbook opens in one step with nothing replayed through commands (and so
 * nothing on the undo stack). Sheet protection, which Univer models differently
 * from Excel, is returned beside it for src/univer/protect.js.
 *
 * Only the sheets the model was read with are built (src/config.js: the
 * typed-in ones, and Lists, hidden, for the dropdowns). A link to a sheet that
 * is not here (the "Start here" crumbs) is dropped with its text.
 */
import { quoteSheet, parseRange, cellName } from "../xlsx/ref.js";

const BORDER = {thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8, mediumDashed: 9,
  mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13};
const HALIGN = {left: 1, center: 2, centerContinuous: 2, right: 3, justify: 4, distributed: 6, fill: 1};
const VALIGN = {top: 1, center: 2, justify: 2, distributed: 2, bottom: 3};

/* Excel's column width (in characters of the default font's widest digit) to pixels, as Excel draws it. */
export function colPx(width, mdw = 7){
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}
const ptPx = pt => Math.round(pt * 96 / 72);

function fontStyle(f, out){
  if(!f) return out;
  if(f.name) out.ff = f.name;
  if(f.size) out.fs = f.size;
  if(f.bold) out.bl = 1;
  if(f.italic) out.it = 1;
  if(f.strike) out.st = {s: 1};
  if(f.underline && f.underline !== "none") out.ul = {s: 1};
  if(f.color) out.cl = {rgb: f.color};
  if(f.vertAlign === "superscript") out.va = 3; else if(f.vertAlign === "subscript") out.va = 2;
  return out;
}
function borderStyle(b){
  if(!b) return null;
  const bd = {};
  const side = (k, s) => { if(b[s]) bd[k] = {s: BORDER[b[s].style] || 1, cl: {rgb: b[s].color}}; };
  side("l", "left"); side("r", "right"); side("t", "top"); side("b", "bottom");
  if(b.diagonal){ if(b.diagDown) bd.tl_br = {s: BORDER[b.diagonal.style] || 1, cl: {rgb: b.diagonal.color}}; if(b.diagUp) bd.bl_tr = {s: BORDER[b.diagonal.style] || 1, cl: {rgb: b.diagonal.color}}; }
  return Object.keys(bd).length ? bd : null;
}
export function xfStyle(xf){
  const s = fontStyle(xf.font, {});
  if(xf.fill) s.bg = {rgb: xf.fill};
  const bd = borderStyle(xf.border);
  if(bd) s.bd = bd;
  if(xf.numFmt && xf.numFmt !== "General") s.n = {pattern: xf.numFmt};
  const a = xf.align || {};
  if(a.h && HALIGN[a.h]) s.ht = HALIGN[a.h];
  s.vt = VALIGN[a.v] || 3;   /* Excel's default is the bottom of the cell */
  if(a.wrap) s.tb = 3;
  if(a.indent) s.pd = {l: 2 + a.indent * 9};
  if(a.rotate){ if(a.rotate === 255) s.tr = {a: 0, v: 1}; else s.tr = {a: a.rotate <= 90 ? a.rotate : -(a.rotate - 90)}; }
  return s;
}
function dxfStyle(d){
  if(!d) return {};
  const s = fontStyle(d.font, {});
  delete s.ff; delete s.fs;   /* a format rule changes colour and weight, not the face or size */
  if(d.fill) s.bg = {rgb: d.fill};
  const bd = borderStyle(d.border);
  if(bd) s.bd = bd;
  return s;
}

export const CHECK_MARGIN = 200;

let uidN = 0;
const uid = p => `${p}${(++uidN).toString(36)}`;

/* Excel serial date to "yyyy-mm-dd", as Univer's date rules take it. */
function serialToIso(n){
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(+n * 86400000));
  return d.toISOString().slice(0, 10);
}

function linkDoc(text, url){
  const t = String(text ?? "");
  return {id: "d", body: {dataStream: t + "\r\n", textRuns: [], paragraphs: [{startIndex: t.length}],
    customRanges: t.length ? [{startIndex: 0, endIndex: t.length - 1, rangeType: 0, rangeId: uid("lnk"), properties: {url}}] : [], customDecorations: []},
    documentStyle: {}};
}

export function buildWorkbook(model, opts = {}){
  const unitId = opts.unitId || "deal-list";
  const styles = {};
  const xfIds = model.styles.xfs.map((xf, i) => { const id = `x${i}`; styles[id] = xfStyle(xf); return id; });
  /* Normal: the style an untouched cell is drawn and typed in. */
  const normal = styles.x0 ? {...styles.x0} : {};
  delete normal.n;
  const font = (model.styles.xfs[0] || {}).font || {};
  const mdw = /calibri/i.test(font.name || "") && (font.size || 11) === 11 ? 7 : Math.max(5, Math.round((font.size || 11) * 0.65));
  const sheetIds = model.sheets.map((s, i) => `sheet${i + 1}`);
  const byName = new Map(model.sheets.map((s, i) => [s.name.toLowerCase(), sheetIds[i]]));
  const sheets = {};
  const extras = {protection: [], unsupported: []};
  const support = new Set((opts.hidden || []).map(n => n.toLowerCase()));
  const dvRes = {}, cfRes = {}, tableRes = {};

  model.sheets.forEach((sh, si) => {
    const id = sheetIds[si];
    const cellData = {};
    for(const [r, row] of sh.cells){
      const out = {};
      for(const [c, cell] of row){
        const d = {};
        if(cell.s) d.s = xfIds[cell.s];
        if(cell.f != null) d.f = "=" + cell.f;
        if(cell.v !== undefined && cell.v !== null){
          if(cell.type === "n"){ d.v = cell.v; d.t = 2; }
          else if(cell.type === "b"){ d.v = cell.v ? 1 : 0; d.t = 3; }
          else { d.v = cell.v; d.t = 1; }
        }
        out[c] = d;
      }
      cellData[r] = out;
    }
    for(const lk of sh.links){
      if(!lk.range) continue;
      let url = lk.target;
      if(lk.location){
        const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(lk.location);
        const sid = m ? byName.get((m[1] || m[2]).replace(/''/g, "'").toLowerCase()) : null;
        url = sid ? `#gid=${sid}&range=${m[3].replace(/\$/g, "")}` : null;
        if(!sid){
          /* a crumb to a sheet this page leaves out: no link, and no text promising one */
          const {startRow: r, startColumn: c} = lk.range;
          if(cellData[r] && cellData[r][c]){ delete cellData[r][c].v; delete cellData[r][c].t; }
          continue;
        }
      }
      if(!url) continue;
      const {startRow: r, startColumn: c} = lk.range;
      cellData[r] = cellData[r] || {};
      const cur = cellData[r][c] || {};
      if(cur.f) continue;   /* a HYPERLINK() formula keeps its formula */
      cur.p = linkDoc(lk.display != null ? lk.display : cur.v, url);
      cellData[r][c] = cur;
    }

    const columnData = {};
    for(const col of sh.cols){
      for(let c = col.min; c <= Math.min(col.max, col.min + 300); c++){
        const cd = {};
        if(col.width != null) cd.w = colPx(col.width, mdw);
        if(col.hidden) cd.hd = 1;
        if(col.style != null && col.style > 0) cd.s = xfIds[col.style];
        columnData[c] = cd;
      }
    }
    const rowData = {};
    for(const [r, rd] of sh.rows){
      const o = {};
      if(rd.ht != null) o.h = ptPx(rd.ht);
      if(rd.hidden) o.hd = 1;
      if(rd.style != null) o.s = xfIds[rd.style];
      rowData[r] = o;
    }

    let maxRow = sh.maxRow, maxCol = sh.maxCol;
    const grow = rg => { if(rg.endRow < 100000 && rg.endRow > maxRow) maxRow = rg.endRow; if(rg.endColumn < 1000 && rg.endColumn > maxCol) maxCol = rg.endColumn; };
    sh.dv.forEach(d => d.ranges.forEach(grow)); sh.cf.forEach(c => c.ranges.forEach(grow)); sh.merges.forEach(grow);
    for(const col of sh.cols) if(col.max < 200 && col.max > maxCol) maxCol = col.max;

    /* Excel's default width: the base width in digits plus 5px of padding, rounded up to a multiple of 8 (8 digits of Calibri 11: 64px) */
    const defW = sh.defaultColWidth != null ? colPx(sh.defaultColWidth, mdw) : Math.ceil(((sh.baseColWidth || 8) * mdw + 5) / 8) * 8;
    sheets[id] = {
      id, name: sh.name, tabColor: sh.tabColor || "", hidden: support.has(sh.name.toLowerCase()) ? 1 : sh.state === "hidden" ? 1 : sh.state === "veryHidden" ? 2 : 0,
      rowCount: Math.max(1000, maxRow + 101), columnCount: Math.max(26, maxCol + 11),
      zoomRatio: (sh.zoom || 100) / 100, defaultColumnWidth: defW, defaultRowHeight: ptPx(sh.defaultRowHeight || 15),
      showGridlines: sh.showGridLines ? 1 : 0, rightToLeft: sh.rightToLeft ? 1 : 0,
      freeze: sh.frozen ? {xSplit: sh.frozen.cols, ySplit: sh.frozen.rows, startRow: sh.frozen.rows, startColumn: sh.frozen.cols} : {xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1},
      mergeData: sh.merges, cellData, rowData, columnData,
      rowHeader: {width: 46}, columnHeader: {height: 20}
    };

    /* data validation: dropdowns, number, date and length rules, with their prompts and messages */
    const dvs = [];
    for(const d of sh.dv){
      const type = {whole: "whole", decimal: "decimal", list: "list", date: "date", time: "time", textLength: "textLength", custom: "custom"}[d.type];
      if(!type) continue;
      let f1 = d.formula1, f2 = d.formula2;
      if(type === "list" && f1 != null){
        if(/^".*"$/.test(f1.trim())) f1 = f1.trim().slice(1, -1);   /* a typed list: a,b,c */
        else f1 = "=" + f1.replace(/^=/, "");
      } else if(type === "date"){
        if(f1 != null && /^-?\d+(\.\d+)?$/.test(f1)) f1 = serialToIso(f1);
        if(f2 != null && /^-?\d+(\.\d+)?$/.test(f2)) f2 = serialToIso(f2);
      } else if(type === "custom" && f1 != null) f1 = "=" + f1.replace(/^=/, "");
      const rule = {uid: uid("dv"), ranges: d.ranges.map(r => ({...r, unitId, sheetId: id})), type, allowBlank: d.allowBlank,
        showDropDown: d.showDropDown, showErrorMessage: d.showErrorMessage, showInputMessage: d.showInputMessage,
        errorStyle: {stop: 1, warning: 2, information: 0}[d.errorStyle] ?? 1};
      if(type !== "list" && type !== "custom") rule.operator = d.operator;
      if(f1 != null) rule.formula1 = f1;
      if(f2 != null) rule.formula2 = f2;
      for(const k of ["error", "errorTitle", "prompt", "promptTitle"]) if(d[k]) rule[k] = d[k];
      if(type === "list") rule.renderMode = 0;   /* as Excel: plain text, the arrow only on the cell selected */
      dvs.push(rule);
    }
    if(dvs.length) dvRes[id] = dvs;

    /* conditional formatting: Excel's priority 1 wins, and so does Univer's first rule.
       CHECKED WHERE THERE IS SOMETHING TO CHECK. Excel works a rule out only for
       the cells on screen; Univer works it out for every cell in its range, and
       the template's checks run to row 2,000 (one counts duplicates across all
       2,000 rows for each cell), which kept the page busy for half a minute.
       So each rule here stops 200 rows below the last row typed in, and is
       stretched as rows are typed further down (src/univer/checks.js). The
       file's own rules are not touched. */
    const cfs = [];
    const cfEnd = Math.max(sh.maxRow, (sh.frozen ? sh.frozen.rows : 0)) + CHECK_MARGIN;
    extras.checkEnd = extras.checkEnd || {};
    extras.checkEnd[id] = cfEnd;
    for(const c of [...sh.cf].sort((a, b) => a.priority - b.priority)){
      const style = dxfStyle(model.styles.dxfs[c.dxf]);
      const ranges = c.ranges.map(r => ({...r, endRow: r.startRow <= cfEnd ? Math.min(r.endRow, cfEnd) : r.endRow, unitId, sheetId: id, fullEndRow: r.endRow}));
      let rule = null;
      if(c.type === "expression" && c.formulas[0] != null) rule = {type: "highlightCell", subType: "formula", value: "=" + c.formulas[0], style};
      else if(c.type === "cellIs"){
        const op = {between: "between", notBetween: "notBetween", equal: "equal", notEqual: "notEqual", greaterThan: "greaterThan",
          greaterThanOrEqual: "greaterThanOrEqual", lessThan: "lessThan", lessThanOrEqual: "lessThanOrEqual"}[c.operator];
        const n = v => (v != null && /^-?\d+(\.\d+)?$/.test(v)) ? +v : null;
        if(op && n(c.formulas[0]) != null)
          rule = {type: "highlightCell", subType: "number", operator: op, value: /between/i.test(op) ? [n(c.formulas[0]), n(c.formulas[1])] : n(c.formulas[0]), style};
        else if(op) rule = {type: "highlightCell", subType: "formula", value: `=${cellName(c.ranges[0].startRow, c.ranges[0].startColumn)}${{equal: "=", notEqual: "<>", greaterThan: ">", greaterThanOrEqual: ">=", lessThan: "<", lessThanOrEqual: "<="}[op] || "="}(${c.formulas[0]})`, style};
      }
      else if(c.type === "containsBlanks") rule = {type: "highlightCell", subType: "text", operator: "containsBlanks", value: "", style};
      else if(c.type === "notContainsBlanks") rule = {type: "highlightCell", subType: "text", operator: "notContainsBlanks", value: "", style};
      else if(c.type === "containsText") rule = {type: "highlightCell", subType: "text", operator: "containsText", value: c.text || "", style};
      else if(c.type === "duplicateValues") rule = {type: "highlightCell", subType: "duplicateValues", style};
      else if(c.type === "uniqueValues") rule = {type: "highlightCell", subType: "uniqueValues", style};
      if(!rule){ extras.unsupported.push(`${sh.name}: a ${c.type} conditional format`); continue; }
      cfs.push({cfId: uid("cf"), ranges, stopIfTrue: !!c.stopIfTrue, rule});
    }
    if(cfs.length) cfRes[id] = cfs;

    /* Excel tables: the name, the range and its filter buttons */
    if(sh.tables.length && opts.tables !== false){
      tableRes[id] = {tables: sh.tables.map(t => ({id: uid("tbl"), name: t.name,
        range: {startRow: t.ref.startRow, endRow: t.ref.endRow, startColumn: t.ref.startColumn, endColumn: t.ref.endColumn},
        options: {showHeader: t.headerRows > 0, showFooter: t.totalsRows > 0, showAutoFilter: t.autoFilter, tableStyleId: "table-default-0"},
        filters: {tableColumnFilterList: []}, columns: []})), tableFilteredOutRows: []};
    }

    if(sh.protection){
      /* the cells a protected sheet still lets you type in: those whose style is unlocked */
      const open = [];
      for(const [r, row] of sh.cells) for(const [c, cell] of row) if(model.styles.xfs[cell.s] && model.styles.xfs[cell.s].locked === false) open.push([r, c]);
      const openCols = sh.cols.filter(c => c.style != null && model.styles.xfs[c.style] && model.styles.xfs[c.style].locked === false).map(c => [c.min, c.max]);
      extras.protection.push({sheetId: id, sheet: sh.name, options: sh.protection, open, openCols});
    }
  });

  /* defined names, workbook-wide or for one sheet; only those that point at sheets that are here */
  const names = {};
  const here = new Set(model.sheets.map(s => s.name.toLowerCase()));
  const refsHere = ref => [...ref.matchAll(/(?:'((?:[^']|'')+)'|([A-Za-z_][\w. ]*))!/g)].every(m => here.has((m[1] || m[2]).replace(/''/g, "'").toLowerCase()));
  for(const n of model.names){
    if(/^_xlnm\./.test(n.name) || !refsHere(n.ref)) continue;
    if(n.local != null && !model.sheets.some(s => s.name === model.allSheetNames[n.local])) continue;
    const nid = uid("dn");
    names[nid] = {id: nid, name: n.name, formulaOrRefString: n.ref.replace(/^=/, ""), localSheetId: n.local == null ? "AllDefaultWorkbook" : sheetIds[model.sheets.findIndex(s => s.name === model.allSheetNames[n.local])], hidden: n.hidden};
  }

  const resources = [
    {name: "SHEET_DEFINED_NAME_PLUGIN", data: JSON.stringify(names)},
    {name: "SHEET_DATA_VALIDATION_PLUGIN", data: JSON.stringify(dvRes)},
    {name: "SHEET_CONDITIONAL_FORMATTING_PLUGIN", data: JSON.stringify(cfRes)},
    {name: "SHEET_TABLE_PLUGIN", data: JSON.stringify(tableRes)}
  ];
  const pick = name => { const i = model.sheets.findIndex(s => s.name === name && !support.has(s.name.toLowerCase())); return i < 0 ? null : sheetIds[i]; };
  const active = pick(model.activeName) || pick(opts.first) || sheetIds[0];
  return {
    snapshot: {id: unitId, name: opts.name || "Deal list", appVersion: "1.0.0", locale: "enUS", styles, defaultStyle: normal,
      sheetOrder: sheetIds, sheets, resources},
    activeSheetId: active, sheetIds, extras
  };
}

export { parseRange, quoteSheet };
