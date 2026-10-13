/* NUMBERS KEPT AS EXCEL KEEPS THEM. Univer reads "4.10%" (typed or pasted) by
   working in binary fractions, and stores 0.040999999999999995 where Excel
   stores 0.041. Excel keeps 15 significant digits; so does this, just before
   any typed or pasted values are written. (Pasted text already becomes numbers
   and dates in Univer itself, day first since boot.js.) */
export function installExcelNumbers(univerAPI){
  const d = univerAPI.addEvent(univerAPI.Event.BeforeCommandExecute, ev => {
    if(ev.id !== "sheet.mutation.set-range-values") return;
    const cv = (ev.params || {}).cellValue || {};
    for(const r of Object.keys(cv)) for(const c of Object.keys(cv[r])){
      const cell = cv[r][c];
      if(cell && typeof cell.v === "number" && isFinite(cell.v) && !cell.f){
        const v = +cell.v.toPrecision(15);
        if(v !== cell.v) cell.v = v;
      }
    }
  });
  return () => d.dispose();
}
