'use strict';
const $=id=>document.getElementById(id), timer=new EyeTimer();
let breakDialog=null, dialogAttempt=0, dialogWatchdog=null;
let audio, motionPhase=0,lastFrame=performance.now();
const defaults={work:'20',rest:'20',exerciseTime:'60',speed:'gentle',increase:false,sound:true,motion:matchMedia('(prefers-reduced-motion: reduce)').matches};
let saved={};try{const current=localStorage.getItem('pro-eye-excel-settings-v12');saved=JSON.parse(current||localStorage.getItem('pro-eye-excel-settings-v1')||'{}');if(!current)saved.sound=true;}catch(e){}
for(const [id,def] of Object.entries(defaults)){const e=$(id),v=id==='rest'?'20':(saved[id]??def);if(e.type==='checkbox')e.checked=!!v;else {e.value=String(v);if(!e.value)e.value=def;}e.addEventListener('change',()=>{save();if(timer.mode==='stopped')render();});}
function save(){const data={};for(const id of Object.keys(defaults))data[id]=$(id).type==='checkbox'?$(id).checked:$(id).value;try{localStorage.setItem('pro-eye-excel-settings-v12',JSON.stringify(data));}catch(e){$('announce').textContent='Settings cannot be saved in this browser; this session still works.';}}
function config(){timer.configure({work:Number($('work').value),rest:20,exercise:Number($('exerciseTime').value)});}
async function enableAudio(){
 if(!$('sound').checked)return false;
 try{audio=audio||new(window.AudioContext||window.webkitAudioContext)();await audio.resume();if(audio.state!=='running')throw Error('Audio is suspended');return true;}
 catch(e){$('audioStatus').textContent='Sound is blocked. Click Test sound and check your browser/tab volume.';return false;}
}
async function chime(){
 if(!await enableAudio())return;
 try{for(const [i,f] of [660,880,660].entries()){const o=audio.createOscillator(),g=audio.createGain();o.connect(g);g.connect(audio.destination);o.frequency.value=f;const at=audio.currentTime+i*0.3;g.gain.setValueAtTime(0.0001,at);g.gain.exponentialRampToValueAtTime(0.18,at+0.025);g.gain.exponentialRampToValueAtTime(0.001,at+0.25);o.start(at);o.stop(at+0.28);} $('audioStatus').textContent='Chime sent to your audio output. Keep your tab and device unmuted.';}
 catch(e){$('audioStatus').textContent='Sound failed. Click Test sound to try again.';}
}
$('testSound').onclick=()=>{ $('sound').checked=true;save();void chime(); };
$('testAlert').onclick=()=>{if(timer.mode!=='stopped')return;config();void enableAudio();timer.enter('rest',Date.now());render();void showBreak();};
$('start').onclick=()=>{config();enableAudio();timer.start(Date.now());$('announce').textContent=backgroundReady?'Timer running. You may close this panel while this workbook stays open.':'Timer running in panel-only mode. Keep the panel open.';render();};
$('stop').onclick=()=>{timer.stop();$('announce').textContent='Stopped. Press Start timer to begin again.';render();};
$('pause').onclick=()=>{timer.pause(Date.now());$('announce').textContent=timer.paused?'Paused.':'Resumed.';render();};
$('exercise').onclick=()=>{config();enableAudio();timer.exercise(Date.now());motionPhase=0;$('announce').textContent='Horizontal exercise started. Follow gently; stop if uncomfortable.';render();};
$('skip').onclick=()=>{timer.start(Date.now());$('announce').textContent='Exercises skipped. A new work timer has started.';render();};
$('sound').onchange=enableAudio;
const labels={stopped:'Ready when you are',work:'Focus time',rest:'Look far',ready:'Break finished',horizontal:'Left ↔ Right',vertical:'Up ↕ Down'};
const instructions={stopped:'Start a focus session. We’ll remind you to look away.',work:'Settle into your work. Your next break is counting down.',rest:'Look away from the screen toward a distant object. Blink naturally.',ready:'Ready for gentle eye movement? Start exercises or skip to your next work session.',horizontal:'Follow the ball gently from left to right.',vertical:'Follow the ball gently up and down.'};
function render(){if(timer.mode!=='rest'&&typeof breakDialog!=='undefined')closeBreakDialog();$('breakAlert').hidden=timer.mode!=='rest';$('testAlert').disabled=timer.mode!=='stopped';const active=timer.mode!=='stopped',exercise=['horizontal','vertical'].includes(timer.mode);const seconds=Math.ceil((active?timer.left:Number($('work').value)*60000)/1000);$('clock').textContent=String(Math.floor(seconds/60)).padStart(2,'0')+':'+String(seconds%60).padStart(2,'0');$('phase').textContent=labels[timer.mode]+(timer.paused&&timer.mode!=='ready'?' · paused':'');$('instruction').textContent=instructions[timer.mode];$('arena').hidden=!exercise;$('progress').style.width=active&&timer.total?Math.max(0,Math.min(100,100*(1-timer.left/timer.total)))+'%':'0%';$('start').disabled=active;$('stop').disabled=!active;$('pause').disabled=!active||timer.mode==='ready';$('pause').textContent=timer.paused?'Resume':'Pause';$('settings').disabled=active;$('exercise').hidden=exercise||timer.mode==='rest';$('skip').hidden=!['ready','horizontal','vertical','rest'].includes(timer.mode);$('cycles').textContent=timer.cycles+' exercise rounds this session';}

function closeBreakDialog(){dialogAttempt++;clearTimeout(dialogWatchdog);dialogWatchdog=null;const d=breakDialog;breakDialog=null;if(d){try{d.close();}catch(e){}}}
async function dialogFallback(reason){
 timer.enter('rest',Date.now());render();void chime();
 $('announce').textContent='Popup unavailable: '+reason+'. Showing the panel reminder instead.';
 if(backgroundReady){try{await Office.addin.showAsTaskpane();}catch(e){}}
}
function showBreak(){
 closeBreakDialog();
 if(!window.Office || !Office.context.ui || !Office.context.ui.displayDialogAsync){void dialogFallback('Open this add-in inside Excel');return;}
 const attempt=dialogAttempt;
 const token=String(Date.now())+'-'+Math.random().toString(36).slice(2);
 const url=new URL('break.html?v=1.3.0',location.href);url.searchParams.set('token',token);
 timer.paused=true;render();
 function fail(reason){if(attempt!==dialogAttempt)return;closeBreakDialog();void dialogFallback(reason);}
 dialogWatchdog=setTimeout(()=>fail('The dialog did not finish loading'),15000);
 try{Office.context.ui.displayDialogAsync(url.href,{height:40,width:35,displayInIframe:true},result=>{
  if(attempt!==dialogAttempt||timer.mode!=='rest'){if(result.status===Office.AsyncResultStatus.Succeeded)result.value.close();return;}
  if(result.status!==Office.AsyncResultStatus.Succeeded){fail(result.error?.message||'Excel could not open the dialog');return;}
  breakDialog=result.value;let started=false;
  breakDialog.addEventHandler(Office.EventType.DialogMessageReceived,arg=>{
   if(attempt!==dialogAttempt || (arg.origin && arg.origin!==location.origin))return;
   let msg;try{msg=JSON.parse(arg.message);}catch(e){return;}
   if(msg.token!==token)return;
   if(msg.type==='ready'&&!started){started=true;clearTimeout(dialogWatchdog);timer.enter('rest',Date.now());render();void chime();}
   if(msg.type==='done'&&started){closeBreakDialog();if(timer.mode==='rest'){timer.start(Date.now());$('announce').textContent='Break finished. Work timer restarted.';render();}}
  });
  breakDialog.addEventHandler(Office.EventType.DialogEventReceived,arg=>{
   if(attempt!==dialogAttempt)return;
   if(!started&&arg.error!==12006){fail('Excel dialog error '+arg.error);return;}
   closeBreakDialog();if(timer.mode==='rest'){timer.start(Date.now());$('announce').textContent='Popup closed. Work timer restarted.';render();}
  });
 });}catch(e){fail(e.message);}
}
function tick(){const before=timer.mode;if(timer.tick(Date.now())){
 if(timer.mode==='rest'){$('announce').textContent='Look far for 20 seconds.';void showBreak();}
 else if(before==='rest'){closeBreakDialog();$('announce').textContent='20-second break finished. The next work timer has started.';chime();}
 else $('announce').textContent=timer.mode==='vertical'?'Now follow up and down.':'Exercise round complete. Work timer restarted.';
 motionPhase=0;
}render();}
setInterval(tick,250);
function animate(now){const dt=Math.min(0.1,(now-lastFrame)/1000);lastFrame=now;if(['horizontal','vertical'].includes(timer.mode)&&!timer.paused){const elapsed=(timer.total-timer.left)/1000;let rate={gentle:0.15,medium:0.25,fast:0.4}[$('speed').value];if($('increase').checked)rate=Math.min(0.75,rate*Math.pow(1.15,Math.floor(elapsed/12)));motionPhase+=dt*rate*2*Math.PI;const pos=$('motion').checked?50:50+Math.sin(motionPhase)*38;$('ball').style.left=(timer.mode==='horizontal'?pos:50)+'%';$('ball').style.top=(timer.mode==='vertical'?pos:50)+'%';}requestAnimationFrame(animate);}requestAnimationFrame(animate);
let backgroundReady=false;
function pauseHiddenExercise(){if(['horizontal','vertical'].includes(timer.mode)&&!timer.paused){timer.pause(Date.now());$('announce').textContent='Exercise paused while hidden. Resume when ready.';render();}}
$('hide').onclick=async()=>{if(!backgroundReady)return;try{await Office.addin.hide();}catch(e){$('announce').textContent='Could not hide the panel: '+e.message;}};
if(window.Office)Office.onReady(async info=>{
 if(info.host!==Office.HostType.Excel){$('host').textContent='Browser preview · keep this page open';return;}
 try {
  if(!Office.context.requirements.isSetSupported('SharedRuntime','1.1'))throw Error('Shared runtime is unavailable in this Excel version.');
  await Office.addin.onVisibilityModeChanged(args=>{
   if(args.visibilityMode===Office.VisibilityMode.hidden)pauseHiddenExercise();
   else render();
  });
  backgroundReady=true;$('hide').disabled=false;
  $('host').textContent='v1.3 · background ready';
  $('announce').textContent='Press Start timer. You can then close the panel; a popup appears for each 20-second break.';
 }catch(e){$('host').textContent='Panel-only mode';$('announce').textContent='Background mode unavailable. Install the v1.3 manifest and reopen the workbook. '+e.message;}
});
document.addEventListener('visibilitychange',()=>{if(document.hidden)pauseHiddenExercise();});
render();
