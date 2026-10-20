/* Pure timer state: no workbook access and no network calls. */
(function(root){
class EyeTimer {
 constructor(){this.mode='stopped';this.paused=false;this.settings={work:20,rest:60,exercise:60};this.cycles=0;this.left=0;this.end=0;}
 configure(s){for(const [k,min,max] of [['work',1,120],['rest',20,300],['exercise',15,180]])if(!Number.isInteger(s[k])||s[k]<min||s[k]>max)throw Error('Invalid '+k);this.settings={...s};}
 duration(mode){return mode==='work'?this.settings.work*60:mode==='rest'?this.settings.rest:this.settings.exercise;}
 enter(mode,now){this.mode=mode;this.paused=false;this.total=this.duration(mode)*1000;this.left=this.total;this.end=now+this.total;}
 start(now){this.enter('work',now);}
 exercise(now){this.enter('horizontal',now);}
 stop(){this.mode='stopped';this.paused=false;this.left=0;}
 pause(now){if(this.mode==='stopped')return;if(this.paused){this.end=now+this.left;this.paused=false;}else{this.left=Math.max(0,this.end-now);this.paused=true;}}
 tick(now){if(this.mode==='stopped'||this.paused)return false;this.left=Math.max(0,this.end-now);if(this.left>0)return false;
 // Do not pretend an unattended exercise was completed after device sleep.
 if(this.mode==='work')this.enter('rest',now);else if(this.mode==='rest'){this.mode='ready';this.left=0;this.paused=true;}else if(this.mode==='horizontal')this.enter('vertical',now);else if(this.mode==='vertical'){this.cycles++;this.enter('work',now);}return true;}
}
if(typeof module!=='undefined')module.exports=EyeTimer;else root.EyeTimer=EyeTimer;
})(typeof window!=='undefined'?window:this);
