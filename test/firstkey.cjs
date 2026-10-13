const {chromium}=require('/opt/node22/lib/node_modules/playwright');
(async()=>{
  const b=await chromium.launch();
  for(const d of [20,30,50,80]){
    const p=await b.newPage({viewport:{width:1440,height:900}});
    await p.goto(process.argv[2]||'http://127.0.0.1:8302/',{waitUntil:'networkidle'});
    await p.waitForFunction(()=>document.getElementById('status').dataset.kind,null,{timeout:60000});
    await p.waitForTimeout(1500);
    await p.evaluate(()=>{ const sh=window.univerAPI.getActiveWorkbook().getSheetByName('Term investments'); sh.activate(); sh.getRange('A19').activate(); });
    await p.waitForTimeout(300); await p.keyboard.type('Fixed deposit',{delay:d}); await p.keyboard.press('Enter'); await p.waitForTimeout(400);
    console.log('first edit at', d, 'ms a key ->', JSON.stringify(await p.evaluate(()=>window.univerAPI.getActiveWorkbook().getSheetByName('Term investments').getRange('A19').getValue())));
    await p.close();
  }
  await b.close();
})();
