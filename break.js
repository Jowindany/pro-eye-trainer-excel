const token=new URLSearchParams(location.search).get('token');
Office.onReady(()=>{
 const end=Date.now()+20000;let finished=false;
 const send=type=>{try{Office.context.ui.messageParent(JSON.stringify({type,token}));}catch(e){document.getElementById('status').textContent='Communication failed. Close this popup with its X button.';}};
 send('ready');
 // Retry readiness in case the parent was still registering its message handler.
 const retry=setInterval(()=>{if(Date.now()<end-18000)send('ready');else clearInterval(retry);},300);
 const tick=()=>{const remaining=Math.max(0,Math.ceil((end-Date.now())/1000));document.getElementById('seconds').textContent=remaining;document.getElementById('progress').value=remaining;if(!remaining&&!finished){finished=true;send('done');document.getElementById('status').textContent='Break complete. Closing…';clearInterval(interval);}};
 const interval=setInterval(tick,100);tick();
});
