// Types into the sheets as a person would and checks what is stored.   node test/typing.cjs [url] [outdir]
const {chromium}=require('/opt/node22/lib/node_modules/playwright');
const URL=process.argv[2]||'http://127.0.0.1:8302/', OUT=process.argv[3]||'/tmp';
let PASS=0, FAIL=0; const ok=(c,m,d)=>{ if(c){ PASS++; console.log('  ok   '+m); } else { FAIL++; console.log('  FAIL '+m+(d!=null?'  :: '+d:'')); } };
(async()=>{
  const b=await chromium.launch(); const p=await b.newPage({viewport:{width:1440,height:900}});
  const errs=[]; p.on('pageerror',e=>errs.push(e.message));
  await p.goto(URL,{waitUntil:'networkidle'});
  await p.waitForFunction(()=>document.getElementById('status').dataset.kind,null,{timeout:60000});
  const calc=await p.evaluate(()=>new Promise(res=>{ const f=window.univerAPI.getFormula(); const t=performance.now(); f.calculationEnd(()=>res(Math.round(performance.now()-t))); setTimeout(()=>res('none within 90s'),90000); }));
  console.log('background calculation after opening (ms):', calc);
  const S=f=>p.evaluate(f);
  const cell=(sheet,a1)=>p.evaluate(([s,a])=>{ const r=window.univerAPI.getActiveWorkbook().getSheetByName(s).getRange(a); const d=window.univerAPI.getActiveWorkbook().save().sheets; const id=window.univerAPI.getActiveWorkbook().getSheetByName(s).getSheetId(); const rc=r.getRange(); const raw=((d[id].cellData||{})[rc.startRow]||{})[rc.startColumn]||null; return {raw, shown:r.getDisplayValue ? r.getDisplayValue() : null}; },[sheet,a1]);
  // select a cell by name box, then type into it like a person
  const typeInto=async(sheet,a1,text)=>{
    await p.evaluate(([s,a])=>{ const sh=window.univerAPI.getActiveWorkbook().getSheetByName(s); sh.activate(); sh.getRange(a).activate(); },[sheet,a1]);
    await p.waitForTimeout(250);
    await p.keyboard.type(text, {delay: 15}); await p.keyboard.press('Enter'); await p.waitForTimeout(400);
  };
  await typeInto('Term investments','E19','01/04/2026');
  let c=await cell('Term investments','E19');
  ok(c.raw && c.raw.v===46113, 'a UK date typed as 01/04/2026 is 1 April 2026 (serial 46113)', JSON.stringify(c));
  ok(c.shown==='01/04/2026', 'and shows as 01/04/2026', c.shown);
  await typeInto('Term investments','F19','30/09/2026');
  c=await cell('Term investments','F19'); ok(c.raw && c.raw.v===46295, '30/09/2026 is a date, not text', JSON.stringify(c.raw));
  await typeInto('Term investments','G19','4.1%');
  c=await cell('Term investments','G19'); ok(c.raw && c.raw.v===0.041, 'a rate typed as 4.1% is 0.041, exactly', JSON.stringify(c.raw)); ok(c.shown==='4.10%', 'and shows as 4.10%', c.shown);
  await typeInto('Term investments','D19','5,000,000');
  c=await cell('Term investments','D19'); ok(c.raw && c.raw.v===5000000, 'an amount typed as 5,000,000 is the number 5000000', JSON.stringify(c.raw)); ok(c.shown==='5,000,000', 'and shows as 5,000,000', c.shown);
  await typeInto('Term investments','B19','Lloyds Bank');
  c=await cell('Term investments','B19'); ok(c.raw && c.raw.v==='Lloyds Bank', 'text stays text', JSON.stringify(c.raw));
  const st=await p.evaluate(async()=>{ const sh=window.univerAPI.getActiveWorkbook().getSheetByName('Term investments'); return {E4: await sh.getRange('E4').getValidatorStatus(), E19: await sh.getRange('E19').getValidatorStatus()}; });
  ok(JSON.stringify(st.E4)==='[["valid"]]' && JSON.stringify(st.E19)==='[["valid"]]', 'the date rule passes the example dates and the typed one', JSON.stringify(st));
  // protected Settings: the version cell refuses, the As at cell takes a date
  await typeInto('Settings','B6','9');
  c=await cell('Settings','B6'); ok(c.raw && c.raw.v===3, 'Settings: the locked template version cannot be typed over', JSON.stringify(c.raw));
  await typeInto('Settings','B4','31/10/2026');
  c=await cell('Settings','B4'); ok(c.raw && c.raw.v===46326, 'Settings: the As at date (unlocked) takes 31/10/2026', JSON.stringify(c.raw));
  await p.screenshot({path:OUT+'/typing.png'});
  console.log('page errors:', JSON.stringify(errs.slice(0,5)));
  console.log(`typing — ${PASS} passed, ${FAIL} failed`);
  await b.close(); process.exit(FAIL?1:0);
})().catch(e=>{ console.log('CRASH', e); process.exit(1); });
