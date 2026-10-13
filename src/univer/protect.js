/* SHEET PROTECTION, AS EXCEL HAS IT.
 *
 * Excel protects a sheet but leaves the cells whose style is "unlocked" open
 * to typing (on Settings: the As at date, the pooling and the period start).
 * Univer's own protection is about who, among collaborators, may edit; it has
 * no unlocked cells. So the workbook's rule is applied here instead: on a
 * protected sheet, typing, pasting or clearing is allowed only where every
 * cell touched is unlocked, and formatting and inserting or deleting rows and
 * columns are refused as Excel refuses them. Widths and heights may change,
 * as the workbook's protection allows.
 */
import { parseRange } from "../xlsx/ref.js";

/* Commands that only look or move: never refused. */
const SAFE = /^sheet\.command\.(select|set-selection|move-selection|expand-selection|scroll|set-scroll|set-worksheet-activate|change-zoom|set-zoom|set-frozen|cancel-frozen|set-.*-frozen|set-selection-frozen|delta-column-width|delta-row-height|set-worksheet-col-width|set-row-height|set-col-auto-width|set-row-is-auto-height|set-col-is-auto-width|set-table-filter|set-table-sort-state)/;
/* Commands that put values in cells: allowed where every cell is unlocked. */
const VALUES = /^sheet\.command\.(set-range-values|clear-selection-content|auto-clear-content|auto-fill|refill|copy-down|copy-right|paste-value|optional-paste)$|paste/;

export function installProtection(univerAPI, built, model){
  const bySheet = new Map();
  for(const p of built.extras.protection){
    const idx = built.sheetIds.indexOf(p.sheetId);
    const sh = model.sheets[idx];
    const open = new Set(p.open.map(([r, c]) => r + "," + c));
    const xfs = model.styles.xfs;
    /* a cell with no cell of its own takes its row's or column's style, or Normal (locked) */
    const locked = (r, c) => {
      if(open.has(r + "," + c)) return false;
      const row = sh.cells.get(r);
      if(row && row.has(c)) return xfs[row.get(c).s] ? xfs[row.get(c).s].locked !== false : true;
      const rd = sh.rows.get(r);
      if(rd && rd.style != null) return xfs[rd.style] ? xfs[rd.style].locked !== false : true;
      const col = sh.cols.find(k => c >= k.min && c <= k.max && k.style != null);
      if(col) return xfs[col.style] ? xfs[col.style].locked !== false : true;
      return true;
    };
    bySheet.set(p.sheetId, {name: p.sheet, locked, options: p.options});
  }
  const lockStructure = !!model.lockStructure;
  if(!bySheet.size && !lockStructure) return () => {};

  const tell = name => univerAPI.showMessage({type: "warning", duration: 4000,
    content: `That cell is locked: ${name} is a protected sheet, and only the cells for typing in can be changed.`});

  const rangesOf = (params, sheet) => {
    const out = [];
    const add = r => { if(r && r.startRow != null) out.push(r); };
    if(params){
      add(params.range); (params.ranges || []).forEach(add);
      if(params.targetRange) add(params.targetRange);
      if(params.cellValue && params.range == null && typeof params.cellValue === "object"){
        for(const r of Object.keys(params.cellValue)) for(const c of Object.keys(params.cellValue[r])) out.push({startRow: +r, endRow: +r, startColumn: +c, endColumn: +c});
      }
    }
    if(!out.length && sheet){
      const sel = sheet.getSelection && sheet.getSelection();
      const list = sel ? (sel.getActiveRangeList ? sel.getActiveRangeList() : [sel.getActiveRange()]) : [];
      for(const fr of list || []) if(fr) out.push(fr.getRange());
    }
    return out;
  };
  const anyLocked = (info, ranges) => {
    for(const r of ranges){
      const rows = Math.min(r.endRow, r.startRow + 5000), cols = Math.min(r.endColumn, r.startColumn + 200);
      for(let i = r.startRow; i <= rows; i++) for(let j = r.startColumn; j <= cols; j++) if(info.locked(i, j)) return true;
    }
    return false;
  };

  const d1 = univerAPI.addEvent(univerAPI.Event.BeforeSheetEditStart, ev => {
    const info = bySheet.get(ev.worksheet.getSheetId());
    if(info && info.locked(ev.row, ev.column)){ ev.cancel = true; tell(info.name); }
  });
  const STRUCTURE = /^sheet\.command\.(insert-sheet|remove-sheet|remove-sheet-confirm|set-worksheet-name|set-worksheet-order|set-worksheet-hidden|set-worksheet-show|copy-sheet)/;
  const d2 = univerAPI.addEvent(univerAPI.Event.BeforeCommandExecute, ev => {
    const id = ev.id || "";
    /* the workbook's structure is locked: no sheet added, removed, renamed, moved, hidden or shown */
    if(lockStructure && STRUCTURE.test(id)){
      ev.cancel = true;
      univerAPI.showMessage({type: "warning", duration: 4000, content: "The workbook's sheets are fixed: the report tool reads them by name and in this order."});
      return;
    }
    if(!/^sheet\.command\.|paste|cut/.test(id) || SAFE.test(id)) return;
    const wb = univerAPI.getActiveWorkbook();
    if(!wb) return;
    const p = ev.params || {};
    const sheet = p.subUnitId ? wb.getSheetBySheetId(p.subUnitId) : wb.getActiveSheet();
    if(!sheet) return;
    const info = bySheet.get(sheet.getSheetId());
    if(!info) return;
    if(VALUES.test(id)){
      if(anyLocked(info, rangesOf(p, sheet))){ ev.cancel = true; tell(info.name); }
      return;
    }
    /* anything else that changes the sheet: formats, merges, rows and columns, rules */
    if(/^sheet\.command\.(set-worksheet-(name|order|hidden|show|tab-color)|set-tab-color|insert-sheet|remove-sheet)/.test(id)) return;
    ev.cancel = true; tell(info.name);
  });
  return () => { d1.dispose(); d2.dispose(); };
}

export { parseRange };
