// Pastes rows copied from Excel (tab-separated text, as Excel puts on the clipboard) into the deal sheets.
const {chromium}=require('/opt/node22/lib/node_modules/playwright');
const URL=process.argv[2]||'http://127.0.0.1:8302/';
/* the stored value, not the formatted one the facade's getValue gives back */
const RAW=`(sheet,a1)=>{ const wb=window.univerAPI.getActiveWorkbook(), sh=wb.getSheetByName(sheet), rg=sh.getRange(a1).getRange(); const c=((wb.save().sheets[sh.getSheetId()].cellData||{})[rg.startRow]||{})[rg.startColumn]; return c ? (c.v===undefined ? null : c.v) : null; }`;
let PASS=0, FAIL=0; const ok=(c,m,d)=>{ if(c){ PASS++; console.log('  ok   '+m); } else { FAIL++; console.log('  FAIL '+m+(d!=null?'  :: '+d:'')); } };
(async()=>{
  const b=await chromium.launch(); const ctx=await b.newContext({viewport:{width:1440,height:900}, permissions:['clipboard-read','clipboard-write']});
  const p=await ctx.newPage(); const errs=[]; p.on('pageerror',e=>errs.push(e.message));
  await p.goto(URL,{waitUntil:'networkidle'});
  await p.waitForFunction(()=>document.getElementById('status').dataset.kind,null,{timeout:60000});
  await p.waitForTimeout(1500);
  const rows=[
    ['Fixed deposit','Coventry Building Society','Building society','7,500,000','01/10/2026','01/04/2027','4.25%','','','A','pasted 1'],
    ['Treasury bill','UK Debt Management Office','UK Government','4,975,000','15/10/2026','14/01/2027','4.10%','5,000,000','','',''],
  ];
  const tsv=rows.map(r=>r.join('\t')).join('\n');
  const html='<table>'+rows.map(r=>'<tr>'+r.map(c=>`<td>${c}</td>`).join('')+'</tr>').join('')+'</table>';
  await p.evaluate(async([tsv,html])=>{ await navigator.clipboard.write([new ClipboardItem({'text/plain':new Blob([tsv],{type:'text/plain'}),'text/html':new Blob([html],{type:'text/html'})})]); },[tsv,html]);
  await p.evaluate(()=>{ const sh=window.univerAPI.getActiveWorkbook().getSheetByName('Term investments'); sh.activate(); sh.getRange('A19').activate(); });
  await p.waitForTimeout(300);
  // click the cell so the grid has the keyboard, then Ctrl+V as a person would
  const box=await p.evaluate(()=>{ const c=document.querySelector('canvas'); const r=c.getBoundingClientRect(); return {x:r.x,y:r.y}; });
  await p.keyboard.press('Control+V'); await p.waitForTimeout(1500);
  const got=await p.evaluate(RAW=>{ const raw=eval(RAW); return [19,20].map(r=>['A','B','C','D','E','F','G','H','J','K'].map(c=>raw('Term investments',c+r))); }, RAW);
  console.log(JSON.stringify(got));
  const [r1,r2]=got;
  ok(r1[0]==='Fixed deposit' && r1[1]==='Coventry Building Society' && r2[1]==='UK Debt Management Office', 'two rows pasted from Excel land in A19:K20', JSON.stringify(got));
  ok(r1[3]===7500000 && r2[7]===5000000, 'amounts arrive as numbers', `${r1[3]} ${r2[7]}`);
  ok(r1[4]===46296 && r1[5]===46478, 'dates arrive as dates, day first (1 Oct 2026, 1 Apr 2027)', `${r1[4]} ${r1[5]}`);
  ok(r1[6]===0.0425 && r2[6]===0.041, 'rates arrive exactly as Excel keeps them: 0.0425 and 0.041', `${r1[6]} ${r2[6]}`);
  // one Undo takes the whole paste back
  await p.keyboard.press('Control+Z'); await p.waitForTimeout(800);
  const after=await p.evaluate(RAW=>{ const raw=eval(RAW); return [raw('Term investments','A19'), raw('Term investments','D19'), raw('Term investments','B20')]; }, RAW);
  ok(after.every(v=>v===null||v===''||v===undefined), 'one Undo takes the whole paste back', JSON.stringify(after));
  // plain text too (a copy from another system rather than Excel)
  await p.evaluate(async()=>{ await navigator.clipboard.writeText('Fixed deposit\tLloyds Bank\tUK bank\t2,000,000\t30/09/2026\t31/12/2026\t4.5%'); });
  await p.evaluate(()=>{ window.univerAPI.getActiveWorkbook().getSheetByName('Term investments').getRange('A21').activate(); });
  await p.keyboard.press('Control+V'); await p.waitForTimeout(1200);
  const t=await p.evaluate(RAW=>{ const raw=eval(RAW); return ['A','B','D','E','F','G'].map(c=>raw('Term investments',c+'21')); }, RAW);
  ok(t[0]==='Fixed deposit' && t[2]===2000000 && t[3]===46295 && t[4]===46387 && Math.abs(t[5]-0.045)<1e-12, 'plain tab-separated text pastes the same way', JSON.stringify(t));
  const shown=await p.evaluate(()=>{ const sh=window.univerAPI.getActiveWorkbook().getSheetByName('Term investments'); return ['D','E','G'].map(c=>sh.getRange(c+'21').getDisplayValue()); });
  ok(shown.join('|')==='2,000,000|30/09/2026|4.50%', 'and shows in the sheet\'s own formats', shown.join('|'));
  console.log('page errors:', JSON.stringify(errs.slice(0,3)));
  console.log(`paste — ${PASS} passed, ${FAIL} failed`);
  await b.close(); process.exit(FAIL?1:0);
})().catch(e=>{ console.log('CRASH', e); process.exit(1); });
