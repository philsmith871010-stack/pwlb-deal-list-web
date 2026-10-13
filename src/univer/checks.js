/* THE CHECKS FOLLOW THE ROWS TYPED. The conditional formats that flag a
   missing field or a figure out of range are built to stop 200 rows below the
   last row in use (src/univer/build.js). When a value is typed within 50 rows
   of that end, every rule on the sheet is stretched another 200 rows, never
   past the range the workbook itself gave it. */
import { CHECK_MARGIN } from "./build.js";

export function installCheckGrowth(univerAPI, built){
  const end = {...(built.extras.checkEnd || {})};
  const grow = (sheet, row) => {
    const id = sheet.getSheetId();
    if(end[id] == null || row < end[id] - 50) return;
    const next = row + CHECK_MARGIN;
    for(const rule of sheet.getConditionalFormattingRules()){
      const ranges = rule.ranges.map(r => ({...r, endRow: Math.max(r.endRow, Math.min(r.fullEndRow ?? r.endRow, next))}));
      if(ranges.some((r, i) => r.endRow !== rule.ranges[i].endRow)) sheet.setConditionalFormattingRule(rule.cfId, {...rule, ranges});
    }
    end[id] = next;
  };
  const d = univerAPI.addEvent(univerAPI.Event.CommandExecuted, ev => {
    if(ev.id !== "sheet.command.set-range-values" && !/paste|auto-fill/.test(ev.id || "")) return;
    const wb = univerAPI.getActiveWorkbook();
    const sheet = wb && wb.getActiveSheet();
    if(!sheet) return;
    const p = ev.params || {};
    let row = p.range ? p.range.endRow : null;
    if(row == null){ const r = sheet.getSelection() && sheet.getSelection().getActiveRange(); row = r ? r.getRange().endRow : null; }
    if(row != null) grow(sheet, row);
  });
  return () => d.dispose();
}
