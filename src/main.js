/* THE DEAL LIST IN THE BROWSER: open the deal list workbook, show its typed-in
   sheets as Excel shows them, let them be filled in, and save the workbook
   back as an .xlsx the report tool reads. */
import { readXlsx } from "./xlsx/read.js";
import { writeXlsx } from "./xlsx/write.js";
import { buildWorkbook } from "./univer/build.js";
import { bootUniver } from "./univer/boot.js";
import { installProtection } from "./univer/protect.js";
import { installCheckGrowth } from "./univer/checks.js";
import { INPUT_SHEETS, SUPPORT_SHEETS, FIRST_SHEET } from "./config.js";

const $ = id => document.getElementById(id);
const status = (t, kind = "") => { const s = $("status"); s.textContent = t; s.dataset.kind = kind; };

const { univerAPI } = bootUniver("sheet");
window.univerAPI = univerAPI;

let CURRENT = null;   /* {model, built, name, unprotect} */

export async function openWorkbook(bytes, name){
  status(`Opening ${name}…`);
  const model = await readXlsx(bytes, { only: [...INPUT_SHEETS, ...SUPPORT_SHEETS] });
  const missing = INPUT_SHEETS.filter(n => !model.sheets.some(s => s.name === n));
  if(missing.length === INPUT_SHEETS.length) throw new Error(`${name} is not a deal list workbook: it has none of the sheets ${INPUT_SHEETS.join(", ")}.`);
  const built = buildWorkbook(model, { name, hidden: SUPPORT_SHEETS, first: FIRST_SHEET });
  if(CURRENT && CURRENT.unprotect) CURRENT.unprotect();
  const old = univerAPI.getActiveWorkbook();
  if(old) univerAPI.disposeUnit(old.getId());
  const wb = univerAPI.createWorkbook(built.snapshot);
  const sh = wb.getSheetBySheetId(built.activeSheetId);
  if(sh) sh.activate();
  /* The cell editor is made the first time it is needed, and keys typed while it
     is being made were lost (only at machine speed, but still): it is opened and
     closed once now, changing nothing. */
  try{ if(wb.startEditing && wb.startEditing()) await wb.endEditingAsync(false); }catch(_){}
  const p1 = installProtection(univerAPI, built, model), p2 = installCheckGrowth(univerAPI, built);
  const unprotect = () => { p1(); p2(); };
  CURRENT = { model, built, name, unprotect };
  window.__deal = CURRENT;
  document.title = `${name} · Deal list`;
  $("fname").textContent = name;
  status(missing.length ? `Opened, but without ${missing.join(", ")}` : "Ready", missing.length ? "bad" : "ok");
  if(built.extras.unsupported.length) console.warn("Not shown:", built.extras.unsupported);
  return CURRENT;
}

/* The workbook as it is now, as .xlsx bytes. */
export async function saveBytes(){
  if(!CURRENT) return null;
  const wb = univerAPI.getActiveWorkbook();
  /* a cell still being typed in is finished first, as Excel does on saving */
  try{ await wb.endEditingAsync(true); }catch(_){}
  return writeXlsx(CURRENT.model, CURRENT.built, wb.save(), { keep: SUPPORT_SHEETS });
}
window.__saveBytes = async () => { const b = await saveBytes(); return b ? Array.from(b) : null; };

async function download(){
  const bytes = await saveBytes();
  if(!bytes) return;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  a.download = CURRENT.name.replace(/(\.xlsx)?$/i, ".xlsx");
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  status(`Saved ${a.download}`, "ok");
}

$("open").addEventListener("click", () => $("file").click());
$("save").addEventListener("click", () => download().catch(err => { console.error(err); status(err.message || String(err), "bad"); }));
$("file").addEventListener("change", async e => {
  const f = e.target.files[0];
  if(!f) return;
  try{ await openWorkbook(new Uint8Array(await f.arrayBuffer()), f.name); }
  catch(err){ console.error(err); status(err.message || String(err), "bad"); }
  e.target.value = "";
});
for(const b of document.querySelectorAll("[data-sample]")){
  b.addEventListener("click", async () => {
    const url = b.dataset.sample;
    try{ const res = await fetch(url); await openWorkbook(new Uint8Array(await res.arrayBuffer()), url.split("/").pop()); }
    catch(err){ console.error(err); status(err.message || String(err), "bad"); }
  });
}

/* The example opens first; ?file=blank opens the blank template instead. */
(async () => {
  const which = new URLSearchParams(location.search).get("file") === "blank" ? "blank-deal-list.xlsx" : "example-deal-list.xlsx";
  try{ const res = await fetch(which); await openWorkbook(new Uint8Array(await res.arrayBuffer()), which); }
  catch(err){ console.error(err); status(err.message || String(err), "bad"); }
})();
