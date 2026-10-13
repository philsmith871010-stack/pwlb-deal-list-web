const {chromium}=require('/opt/node22/lib/node_modules/playwright');
(async()=>{
  const b=await chromium.launch(); const p=await b.newPage({viewport:{width:1440,height:900}});
  await p.addInitScript(()=>{ window.__long=[]; new PerformanceObserver(l=>{ for(const e of l.getEntries()) window.__long.push(Math.round(e.duration)); }).observe({type:'longtask', buffered:true}); });
  const t0=Date.now();
  await p.goto(process.argv[2]||'http://127.0.0.1:8302/',{waitUntil:'networkidle'});
  await p.waitForFunction(()=>document.getElementById('status').dataset.kind,null,{timeout:60000});
  const ready=Date.now()-t0;
  await p.waitForTimeout(8000);
  const l=await p.evaluate(()=>window.__long);
  console.log('ready after', ready,'ms; long tasks:', l.length, 'longest', Math.max(0,...l), 'ms; total blocked', l.reduce((a,b)=>a+b,0),'ms');
  console.log('over 200ms:', JSON.stringify(l.filter(x=>x>200)));
  await b.close();
})();
