'use strict';
const $=id=>document.getElementById(id);
const invoke=(name,args)=>window.__TAURI__.core.invoke(name,args);
const fields=['name','url','token','pattern','selected','day_sheet','night_sheet','startup','auto_update','remote_update','day','night','refresh','tomorrow'];
const bools=new Set(['startup','auto_update','remote_update']);
const hotkeys=['day','night','refresh','tomorrow'];
const titles={general:['一般設定','端末とエージェントの動作を設定します。'],connection:['接続設定','Machine-control への接続情報を管理します。'],excel:['Excel 設定','対象のブックと勤務シートを指定します。'],keys:['ショートカット','キーボード操作をカスタマイズします。'],updates:['更新・管理','バージョンと翌日の予定を管理します。']};
let ready=false,working=false,updating=false,dirty=false,statusBusy=false;
function say(message,error=false){$('message').textContent=String(message);$('message').classList.toggle('error',error)}
function buttons(){for(const id of ['save','restart','stop','checkUpdate','importTomorrow'])$(id).disabled=!ready||working||updating;for(const key of fields)$(key).disabled=!ready||working||updating}
function setDirty(value){dirty=value;$('dirty').textContent=dirty?'未保存の変更があります':'保存済み'}
function showStatus(s){const label=s.busy?'処理中':s.running?'起動中':'停止中';$('agentState').textContent=label;$('sideState').textContent=label;$('agentVersion').textContent=s.version?'v'+s.version:label;$('statusDot').className='dot '+(s.busy?'busy':s.running?'on':'');document.querySelector('.aside-bottom .dot').className=$('statusDot').className}
function form(){return Object.fromEntries(fields.map(key=>[key,bools.has(key)?$(key).checked:$(key).value]))}
async function refreshStatus(){if(statusBusy||working||updating)return;statusBusy=true;try{showStatus(await invoke('agent_action',{action:'status'}))}catch{ $('agentState').textContent='確認できません' }finally{statusBusy=false}}
async function persist(){await invoke('save_settings',{form:form()});setDirty(false)}
async function run(work){if(working||updating||!ready)return;working=true;buttons();try{await work()}catch(e){say(e,true)}finally{working=false;buttons()}}
for(const key of hotkeys){for(const prefix of ['Ctrl+','Ctrl+Shift+'])for(let i=0;i<=9;i++)$(key).add(new Option(prefix+i,prefix+i))}
for(const key of fields)$(key).addEventListener('input',()=>setDirty(true));
for(const item of document.querySelectorAll('.nav'))item.addEventListener('click',()=>{const tab=item.dataset.tab;for(const button of document.querySelectorAll('.nav'))button.classList.toggle('active',button===item);for(const panel of document.querySelectorAll('.panel'))panel.hidden=panel.id!==tab;$('heading').textContent=titles[tab][0];$('subtitle').textContent=titles[tab][1]});
$('showToken').addEventListener('click',()=>{const visible=$('token').type==='password';$('token').type=visible?'text':'password';$('showToken').textContent=visible?'非表示':'表示'});
$('save').addEventListener('click',()=>run(async()=>{await persist();say('設定を保存しました。再起動すると適用されます。')}));
$('restart').addEventListener('click',()=>run(async()=>{await persist();say('再起動しています…');showStatus(await invoke('agent_action',{action:'restart'}));say('エージェントを再起動しました。')}));
$('stop').addEventListener('click',()=>run(async()=>{showStatus(await invoke('agent_action',{action:'stop'}));say('エージェントを停止しました。')}));
$('importTomorrow').addEventListener('click',()=>run(async()=>{showStatus(await invoke('agent_action',{action:'tomorrow'}));say('翌日の予定更新を受け付けました。Excel 下部で進行状況を確認してください。')}));
$('checkUpdate').addEventListener('click',()=>run(async()=>{await persist();await invoke('start_update');updating=true;say('GitHub で更新を確認しています…');buttons()}));
let updatePolling=false;
async function pollUpdate(){
 if(!updating||updatePolling)return;updatePolling=true;
 try{const state=await invoke('update_state');
 const msg={CHECKING:'最新バージョンを確認しています…',DOWNLOADING:'更新ファイルをダウンロードしています…',UPDATE_READY:'更新を適用します。',UP_TO_DATE:'最新版を使用しています。',UPDATE_DEFERRED:'予定更新中のため保留しました。',UPDATE_FAILED:'更新に失敗しました。ログを確認してください。'};
 if(msg[state])say(msg[state],state==='UPDATE_FAILED');
 if(state==='UPDATE_READY'){await invoke('close_after_update');return}
 if(['UP_TO_DATE','UPDATE_DEFERRED','UPDATE_FAILED','IDLE'].includes(state)){updating=false;buttons()}
 }catch(e){say(e,true);updating=false;buttons()}finally{updatePolling=false}
}
async function init(){buttons();try{const result=await invoke('load_settings');$('version').textContent='Settings v'+result.version;for(const key of fields){if(bools.has(key))$(key).checked=!!result.form[key];else $(key).value=result.form[key]??''}ready=true;setDirty(false);say('準備完了');buttons();void refreshStatus()}catch(e){say(e,true)}}
void init();setInterval(refreshStatus,15000);setInterval(pollUpdate,1000);
