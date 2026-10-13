/* SAVING: THE USER'S OWN WORKBOOK, WITH THE TYPED-IN SHEETS' CELLS REPLACED.
 *
 * The file that was opened is written back part for part. Only the cells of the
 * sheets this page shows are rewritten (each sheet's <sheetData>, from what is
 * in the grid now), and each Excel table on them is stretched over any rows
 * typed beneath it. Everything else (the other sheets, their formulas and
 * charts, styles, names, the dropdown and checking rules) is the file as it
 * came. Excel is told to work every formula out again when it opens it, so the
 * dashboards follow the new rows.
 *
 * A cell keeps the workbook's own style for its place. Formatting changed in
 * the browser (bold, a fill) is not carried into the file: this page is for
 * the deals and the projection, and the workbook's look is the template's.
 */
import { zip } from "./zip.js";
import { cellName, parseRange } from "./ref.js";

const enc = new TextEncoder(), dec = new TextDecoder();
const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");

function cellText(cell){
  if(cell.p && cell.p.body) return cell.p.body.dataStream.replace(/\r\n$/, "").replace(/\r/g, "\n");
  return cell.v;
}
/* The xf index a Univer style id stands for ("x12" -> 12), or null. */
const xfOf = s => typeof s === "string" && /^x\d+$/.test(s) ? +s.slice(1) : null;

function sheetDataXml(sheetModel, data, rowData){
  /* row attributes as the file had them, then as the grid has them now */
  const rows = new Map();
  for(const [r, rd] of sheetModel.rows) rows.set(r, {...rd});
  for(const [r, rd] of Object.entries(rowData || {})){
    const o = rows.get(+r) || {};
    if(rd.h != null) o.ht = Math.round(rd.h * 0.75 * 100) / 100;
    if(rd.hd) o.hidden = true; else delete o.hidden;
    rows.set(+r, o);
  }
  const cellRows = Object.keys(data || {}).map(Number);
  const all = [...new Set([...rows.keys(), ...cellRows])].sort((a, b) => a - b);
  const colStyle = c => { const k = sheetModel.cols.find(x => c >= x.min && c <= x.max && x.style != null); return k ? k.style : null; };
  let maxR = 0, maxC = 0, out = "<sheetData>";
  for(const r of all){
    const ra = rows.get(r) || {};
    let attrs = ` r="${r + 1}"`;
    if(ra.ht != null) attrs += ` ht="${ra.ht}" customHeight="1"`;
    if(ra.hidden) attrs += ` hidden="1"`;
    if(ra.style != null) attrs += ` s="${ra.style}" customFormat="1"`;
    const cols = data[r] ? Object.keys(data[r]).map(Number).sort((a, b) => a - b) : [];
    let cells = "";
    for(const c of cols){
      const cell = data[r][c];
      if(!cell) continue;
      const orig = (sheetModel.cells.get(r) || new Map()).get(c);
      let s = xfOf(cell.s);
      if(s == null) s = orig ? orig.s : colStyle(c);
      const sAttr = s ? ` s="${s}"` : "";
      const v = cellText(cell);
      const ref = cellName(r, c);
      if(cell.f){
        cells += `<c r="${ref}"${sAttr}><f>${esc(cell.f.replace(/^=/, ""))}</f></c>`;
      } else if(v === undefined || v === null || v === ""){
        if(s) cells += `<c r="${ref}"${sAttr}/>`;
        else continue;
      } else if(cell.t === 3 || typeof v === "boolean"){
        cells += `<c r="${ref}"${sAttr} t="b"><v>${v === true || v === 1 ? 1 : 0}</v></c>`;
      } else if(typeof v === "number" && isFinite(v)){
        cells += `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
      } else {
        const t = String(v);
        cells += `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${esc(t)}</t></is></c>`;
      }
      if(v !== undefined && v !== null && v !== ""){ if(r > maxR) maxR = r; if(c > maxC) maxC = c; }
    }
    if(!cells && !Object.keys(ra).length) continue;
    out += cells ? `<row${attrs}>${cells}</row>` : `<row${attrs}/>`;
  }
  return {xml: out + "</sheetData>", maxR, maxC};
}

/* The last row with anything typed in a table's columns, below its header. */
function lastRowIn(data, t){
  let last = t.ref.startRow + t.headerRows;
  for(const k of Object.keys(data || {}).map(Number)){
    if(k <= t.ref.startRow) continue;
    for(let c = t.ref.startColumn; c <= t.ref.endColumn; c++){
      const cell = data[k][c];
      const v = cell ? cellText(cell) : null;
      if(v !== undefined && v !== null && v !== ""){ if(k > last) last = k; break; }
    }
  }
  return last;
}

/* opts.keep: names of sheets written back exactly as they came (the hidden Lists). */
export async function writeXlsx(model, built, snapshot, opts = {}){
  const parts = new Map(model.parts);
  const keep = new Set((opts.keep || []).map(n => n.toLowerCase()));
  model.sheets.forEach((sh, i) => {
    if(keep.has(sh.name.toLowerCase())) return;
    const id = built.sheetIds[i];
    const ws = snapshot.sheets[id];
    if(!ws) return;
    const xml = dec.decode(parts.get(sh.path));
    const {xml: sd, maxR, maxC} = sheetDataXml(sh, ws.cellData || {}, ws.rowData || {});
    let next = xml.replace(/<sheetData\/>|<sheetData>[\s\S]*<\/sheetData>/, sd);
    next = next.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${cellName(Math.max(maxR, 0), Math.max(maxC, 0))}"/>`);
    parts.set(sh.path, enc.encode(next));
    /* each table stretched over the rows typed beneath it (never shrunk below its own size) */
    for(const t of sh.tables){
      const last = lastRowIn(ws.cellData, t);
      if(last <= t.ref.endRow) continue;
      const ref = `${cellName(t.ref.startRow, t.ref.startColumn)}:${cellName(last, t.ref.endColumn)}`;
      const tx = dec.decode(parts.get(t.path)).replace(/(<table\b[^>]*\bref=")[^"]*(")/, `$1${ref}$2`).replace(/(<autoFilter\b[^>]*\bref=")[^"]*(")/, `$1${ref}$2`);
      parts.set(t.path, enc.encode(tx));
    }
  });
  /* Excel works everything out again on opening */
  let wb = dec.decode(parts.get(model.wbPath));
  if(/<calcPr\b/.test(wb)) wb = wb.replace(/<calcPr\b([^>]*?)\s*(\/?)>/, (m, a, sl) => `<calcPr${a.replace(/\s*fullCalcOnLoad="[^"]*"/, "")} fullCalcOnLoad="1"${sl ? "/" : ""}>`);
  else if(/<\/definedNames>/.test(wb)) wb = wb.replace(/<\/definedNames>/, `</definedNames><calcPr fullCalcOnLoad="1"/>`);   /* calcPr's place: after the names */
  else wb = wb.replace(/<\/sheets>/, `</sheets><calcPr fullCalcOnLoad="1"/>`);
  parts.set(model.wbPath, enc.encode(wb));
  return zip(parts);
}

export { parseRange };
