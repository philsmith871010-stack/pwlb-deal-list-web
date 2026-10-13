// Saves the workbook from the page and reads it back with the PWLB report tool's own reader:
// unchanged, it must read exactly as the original; edited, it must carry the edits.
//   node test/roundtrip.cjs [appUrl] [toolUrl] [outdir]
const {chromium}=require('/opt/node22/lib/node_modules/playwright');
const fs=require('fs');
const APP=process.argv[2]||'http://127.0.0.1:8302/', TOOL=process.argv[3]||'http://127.0.0.1:8201/pwlbreports-tool.html', OUT=process.argv[4]||'/tmp';
let PASS=0, FAIL=0; const ok=(c,m,d)=>{ if(c){ PASS++; console.log('  ok   '+m); } else { FAIL++; console.log('  FAIL '+m+(d!=null?'  :: '+String(d).slice(0,600):'')); } };
(async()=>{
  const b=await chromium.launch();
  const app=await b.newPage({viewport:{width:1440,height:900}}), tool=await (await b.newContext()).newPage();
  await tool.route('**/pwlbreports-skin.js', r=>r.abort());
  const errs=[]; app.on('pageerror',e=>errs.push(e.message));
  await app.goto(APP,{waitUntil:'networkidle'});
  await app.waitForFunction(()=>document.getElementById('status').dataset.kind,null,{timeout:60000});
  await tool.goto(TOOL,{waitUntil:'networkidle'});
  const read=async bytes=>tool.evaluate(async arr=>{ const r=REG.readRegister(await XL.readWorkbook(new Uint8Array(arr)));
    const pick=x=>x.map(o=>{ const c={...o}; delete c.issues; delete c.row; return c; });
    return {asAt:String(r.asAt), hra:r.hra, version:r.version, periodFrom:String(r.periodFrom), term:pick(r.termInvestments), call:pick(r.callAccounts), loans:pick(r.loans), funds:pick(r.pooledFunds),
      projections:r.projections, issues:[...(r.issues||[])]}; }, Array.from(bytes));
  const orig=fs.readFileSync(__dirname+'/../public/example-deal-list.xlsx');
  const A=await read(orig);
  const saved0=await app.evaluate(()=>window.__saveBytes());
  fs.writeFileSync(OUT+'/saved-unchanged.xlsx', Buffer.from(saved0));
  const B=await read(saved0);
  ok(A.term.length===15 && A.loans.length===10, 'the original reads as the example: 15 term deals, 10 loans', `${A.term.length} ${A.loans.length}`);
  for(const k of ['asAt','hra','version','periodFrom','term','call','loans','funds','projections','issues'])
    ok(JSON.stringify(A[k])===JSON.stringify(B[k]), `saved unchanged, the tool reads the same ${k}`, JSON.stringify(A[k]).slice(0,200)+' VS '+JSON.stringify(B[k]).slice(0,200));

  // edits, typed as a person would
  const typeInto=async(sheet,a1,text)=>{
    await app.evaluate(([s,a])=>{ const sh=window.univerAPI.getActiveWorkbook().getSheetByName(s); sh.activate(); sh.getRange(a).activate(); },[sheet,a1]);
    await app.waitForTimeout(200); await app.keyboard.type(text,{delay:10}); await app.keyboard.press('Enter'); await app.waitForTimeout(300);
  };
  const row19=[['A19','Fixed deposit'],['B19','Coventry Building Society'],['C19','Building society'],['D19','7,500,000'],['E19','01/10/2026'],['F19','01/04/2027'],['G19','4.25%']];
  for(const [a,t] of row19) await typeInto('Term investments',a,t);
  await typeInto('Settings','B4','31/10/2026');
  await typeInto('Projections','B7','60.1');
  await typeInto('Loans','D5','12,000,000');
  const saved1=await app.evaluate(()=>window.__saveBytes());
  fs.writeFileSync(OUT+'/saved-edited.xlsx', Buffer.from(saved1));
  const C=await read(saved1);
  const t=C.term[C.term.length-1]||{};
  ok(C.term.length===16, 'the new deal is read: 16 term deals', C.term.length);
  ok(t.counterparty==='Coventry Building Society' && t.principal===7500000 && Math.abs(t.rate-0.0425)<1e-12 && t.type==='Fixed deposit' && t.cpType==='Building society', 'and read as typed: counterparty, type, £7.5m at 4.25%', JSON.stringify(t));
  ok(/2026-10-01|Oct 01 2026|1 Oct/.test(String(t.start)) || (t.start && new Date(t.start).toISOString().slice(0,10)==='2026-10-01'), 'its start date is 1 October 2026 (UK order kept)', String(t.start));
  ok(/2026-10-31|Oct 31 2026/.test(C.asAt) || new Date(C.asAt).toISOString().slice(0,10)==='2026-10-31', 'the As at date is 31 October 2026', C.asAt);
  ok(C.loans[1] && C.loans[1].principal===12000000, 'the second loan is now £12m', C.loans[1] && C.loans[1].principal);
  const gfRow=C.projections && C.projections.gf && C.projections.gf.rows ? C.projections.gf.rows.find(r=>/2026\/27/.test(String(r[0]))) : null;
  ok(JSON.stringify(C.projections).includes('60.1'), 'the projection carries the 60.1 typed for 2026/27', JSON.stringify(gfRow||C.projections).slice(0,300));
  ok(C.issues.length===A.issues.length, 'no new problems reported by the tool', JSON.stringify(C.issues));
  console.log('page errors:', JSON.stringify(errs.slice(0,5)));
  console.log(`round trip — ${PASS} passed, ${FAIL} failed`);
  await b.close(); process.exit(FAIL?1:0);
})().catch(e=>{ console.log('CRASH', e); process.exit(1); });
