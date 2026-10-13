// Screenshots each visible sheet of the app.   node test/shots.cjs [url] [outdir]
const {chromium}=require('/opt/node22/lib/node_modules/playwright');
const URL=process.argv[2]||'http://127.0.0.1:8302/', OUT=process.argv[3]||'/tmp';
(async()=>{
  const b=await chromium.launch(); const p=await b.newPage({viewport:{width:1440,height:900}});
  const errs=[], outside=[]; p.on('pageerror',e=>errs.push(e.message)); p.on('console',m=>{ if(m.type()==='error'||m.type()==='warning') errs.push(m.type()+': '+m.text().slice(0,200)); });
  p.on('request',r=>{ const u=r.url(); if(!/^http:\/\/127\.0\.0\.1/.test(u) && !/^(data|blob):/.test(u)) outside.push(u); });
  await p.goto(URL,{waitUntil:'networkidle'});
  await p.waitForFunction(()=>document.getElementById('status').dataset.kind,null,{timeout:60000});
  console.log('status:', await p.textContent('#status'));
  const names=await p.evaluate(()=>window.univerAPI.getActiveWorkbook().getSheets().filter(s=>!s.isSheetHidden()).map(s=>s.getSheetName()));
  console.log('tabs:', names.join(' | '));
  for(const n of names){
    await p.evaluate(n=>window.univerAPI.getActiveWorkbook().getSheetByName(n).activate(), n);
    await p.waitForTimeout(700);
    await p.screenshot({path:`${OUT}/sheet-${n.replace(/\W+/g,'-')}.png`});
  }
  console.log('messages:', JSON.stringify(errs.slice(0,10)));
  console.log('outside requests:', JSON.stringify(outside));
  await b.close();
})().catch(e=>{ console.log('CRASH', e); process.exit(1); });
