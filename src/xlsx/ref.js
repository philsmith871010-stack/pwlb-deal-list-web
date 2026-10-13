/* A1 references: letters and numbers to row and column (both from 0) and back. */

export function colToNum(s){
  let n = 0;
  for(let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

export function numToCol(n){
  let s = "";
  n += 1;
  while(n > 0){ const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/* "B12" -> {r: 11, c: 1} */
export function parseCell(a){
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(a);
  if(!m) return null;
  return {r: +m[2] - 1, c: colToNum(m[1])};
}

/* "A4:K18" -> {startRow, startColumn, endRow, endColumn}; a single cell is a range of one.
   Whole columns ("A:A") and whole rows ("4:4") run to the sheet's limits. */
export function parseRange(a, maxRow = 1048575, maxCol = 16383){
  const s = a.replace(/\$/g, "");
  const [x, y] = s.split(":");
  const col = /^([A-Z]{1,3})$/, row = /^(\d+)$/;
  if(y !== undefined && col.test(x) && col.test(y))
    return {startRow: 0, endRow: maxRow, startColumn: colToNum(x), endColumn: colToNum(y)};
  if(y !== undefined && row.test(x) && row.test(y))
    return {startRow: +x - 1, endRow: +y - 1, startColumn: 0, endColumn: maxCol};
  const a1 = parseCell(x), b1 = parseCell(y === undefined ? x : y);
  if(!a1 || !b1) return null;
  return {startRow: Math.min(a1.r, b1.r), endRow: Math.max(a1.r, b1.r), startColumn: Math.min(a1.c, b1.c), endColumn: Math.max(a1.c, b1.c)};
}

export const cellName = (r, c) => numToCol(c) + (r + 1);

/* A sheet name as a formula would quote it. */
export function quoteSheet(name){
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/* MOVING A FORMULA (Excel's shared formulas): every relative reference shifts by
   the rows and columns between the cell that holds the formula and the one it is
   shared with; absolute ($) parts stay. Text in double quotes and sheet names in
   single quotes are left alone, and so are names and functions (letters followed
   by a bracket, or with an underscore, are never a cell). */
export function shiftFormula(f, dr, dc){
  if(!dr && !dc) return f;
  let out = "", i = 0;
  const re = /(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])/y;
  const colOnly = /(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})(?![A-Za-z0-9_(])/y;
  const rowOnly = /(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_(])/y;
  while(i < f.length){
    const ch = f[i];
    if(ch === '"'){ const j = f.indexOf('"', i + 1) < 0 ? f.length : skipQuoted(f, i, '"'); out += f.slice(i, j); i = j; continue; }
    if(ch === "'"){ const j = skipQuoted(f, i, "'"); out += f.slice(i, j); i = j; continue; }
    const prev = i ? f[i - 1] : "";
    if(/[A-Za-z0-9_.\]]/.test(prev)){ out += ch; i++; continue; }
    colOnly.lastIndex = i;
    let m = colOnly.exec(f);
    if(m){
      const c1 = m[1] ? m[2] : numToColSafe(colToNum(m[2]) + dc), c2 = m[3] ? m[4] : numToColSafe(colToNum(m[4]) + dc);
      out += `${m[1]}${c1}:${m[3]}${c2}`; i = colOnly.lastIndex; continue;
    }
    rowOnly.lastIndex = i;
    m = rowOnly.exec(f);
    if(m && (prev === "" || /[(,:!=+\-*/^&<> ]/.test(prev))){
      const r1 = m[1] ? m[2] : String(+m[2] + dr), r2 = m[3] ? m[4] : String(+m[4] + dr);
      out += `${m[1]}${r1}:${m[3]}${r2}`; i = rowOnly.lastIndex; continue;
    }
    re.lastIndex = i;
    m = re.exec(f);
    if(m){
      const col = m[1] ? m[2] : numToColSafe(colToNum(m[2]) + dc);
      const row = m[3] ? m[4] : String(+m[4] + dr);
      out += `${m[1]}${col}${m[3]}${row}`; i = re.lastIndex; continue;
    }
    out += ch; i++;
  }
  return out;
}
function numToColSafe(n){ return n < 0 ? "#REF!" : numToCol(n); }
function skipQuoted(f, i, q){
  let j = i + 1;
  while(j < f.length){
    if(f[j] === q){ if(f[j + 1] === q){ j += 2; continue; } return j + 1; }
    j++;
  }
  return f.length;
}
