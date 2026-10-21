'use strict';
const $=id=>document.getElementById(id), timer=new EyeTimer();
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
function render(){$('breakAlert').hidden=timer.mode!=='rest';$('testAlert').disabled=timer.mode!=='stopped';const active=timer.mode!=='stopped',exercise=['horizontal','vertical'].includes(timer.mode);const seconds=Math.ceil((active?timer.left:Number($('work').value)*60000)/1000);$('clock').textContent=String(Math.floor(seconds/60)).padStart(2,'0')+':'+String(seconds%60).padStart(2,'0');$('phase').textContent=labels[timer.mode]+(timer.paused&&timer.mode!=='ready'?' · paused':'');$('instruction').textContent=instructions[timer.mode];$('arena').hidden=!exercise;$('progress').style.width=active&&timer.total?Math.max(0,Math.min(100,100*(1-timer.left/timer.total)))+'%':'0%';$('start').disabled=active;$('stop').disabled=!active;$('pause').disabled=!active||timer.mode==='ready';$('pause').textContent=timer.paused?'Resume':'Pause';$('settings').disabled=active;$('exercise').hidden=exercise||timer.mode==='rest';$('skip').hidden=!['ready','horizontal','vertical','rest'].includes(timer.mode);$('cycles').textContent=timer.cycles+' exercise rounds this session';}
async function showBreak() {
 if (!backgroundReady) {void chime();return;}
 const expectedEnd=timer.end;
 try {
  await Office.addin.showAsTaskpane();
  if(timer.mode==='rest'){void chime();$('breakAlert').scrollIntoView({block:'start',behavior:'instant'});}
  // Give the user a full 20 seconds after the reminder is displayed.
  if(timer.mode==='rest' && !timer.paused && timer.end===expectedEnd){timer.enter('rest',Date.now());render();}
 } catch(e) {
  $('announce').textContent='Break due. Excel could not reopen the panel: '+e.message;
 }
}
function tick(){const before=timer.mode;if(timer.tick(Date.now())){
 if(timer.mode==='rest'){$('announce').textContent='Look far for 20 seconds.';void showBreak();}
 else if(before==='rest'){$('announce').textContent='20-second break finished. The next work timer has started.';chime();}
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
  $('host').textContent='v1.2 · background ready';
  $('announce').textContent='Press Start timer. You can then close the panel; it reopens for each 20-second break.';
 }catch(e){$('host').textContent='Panel-only mode';$('announce').textContent='Background mode unavailable. Install the v1.2 manifest and reopen the workbook. '+e.message;}
});
document.addEventListener('visibilitychange',()=>{if(document.hidden)pauseHiddenExercise();});
render();
