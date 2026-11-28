#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::{env, fs, io::Write, os::windows::process::CommandExt, path::{Path,PathBuf}, process::{Child,Command,Stdio}, sync::Mutex, time::{SystemTime,UNIX_EPOCH}};
use winreg::{enums::HKEY_CURRENT_USER,RegKey};
#[path="../../src/hotkeys.rs"]
mod hotkeys;
const VERSION:&str=env!("CARGO_PKG_VERSION");
const NO_WINDOW:u32=0x08000000;
const DEFAULT_CFG:&str=include_str!("../../settings.json");
const LEGACY:&str=include_str!("../../scripts/settings_gui.ps1");
const UPDATE_STATE:&str=include_str!("../../scripts/settings_update_state.ps1");
const BRIDGE:&str=include_str!("agent-bridge.ps1");

#[derive(Debug,Clone,Serialize,Deserialize)]
struct SettingsForm {
    name:String, url:String, #[serde(default)] token:String, #[serde(default)] token_configured:bool,
    pattern:String, selected:String, day_sheet:String, night_sheet:String,
    startup:bool, auto_update:bool, remote_update:bool,
    day:String, night:String, refresh:String, tomorrow:String,
    auto_save:bool, auto_save_minutes:Option<i64>,
}
#[cfg(windows)]
mod dpapi {
    use std::{ffi::c_void,ptr};
    #[repr(C)]
    struct DataBlob { cb_data:u32, pb_data:*mut u8 }
    #[link(name="crypt32")]
    extern "system" {
        fn CryptProtectData(data:*const DataBlob,desc:*const u16,entropy:*const DataBlob,reserved:*mut c_void,prompt:*const c_void,flags:u32,out:*mut DataBlob)->i32;
    }
    #[link(name="kernel32")]
    extern "system" { fn LocalFree(mem:*mut c_void)->*mut c_void; }
    fn b64(bytes:&[u8])->String{
        const T:&[u8]=b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        let mut out=String::with_capacity(bytes.len().div_ceil(3)*4);
        for chunk in bytes.chunks(3){
            let b0=chunk[0] as u32;
            let b1=chunk.get(1).copied().unwrap_or(0) as u32;
            let b2=chunk.get(2).copied().unwrap_or(0) as u32;
            let n=(b0<<16)|(b1<<8)|b2;
            out.push(T[((n>>18)&63) as usize] as char);
            out.push(T[((n>>12)&63) as usize] as char);
            out.push(if chunk.len()>1{T[((n>>6)&63) as usize] as char}else{'='});
            out.push(if chunk.len()>2{T[(n&63) as usize] as char}else{'='});
        }
        out
    }
    pub fn protect(raw:&str)->Result<String,String>{
        let trimmed=raw.trim();
        if trimmed.is_empty()||trimmed.starts_with("dpapi:"){return Ok(trimmed.to_owned())}
        let bytes=trimmed.as_bytes();
        let input=DataBlob{cb_data:bytes.len() as u32,pb_data:bytes.as_ptr() as *mut u8};
        let mut out=DataBlob{cb_data:0,pb_data:ptr::null_mut()};
        let ok=unsafe{CryptProtectData(&input,ptr::null(),ptr::null(),ptr::null_mut(),ptr::null(),0x1,&mut out)};
        if ok==0{return Err("認証キーを暗号化できませんでした。".into())}
        let enc=unsafe{std::slice::from_raw_parts(out.pb_data,out.cb_data as usize).to_vec()};
        unsafe{LocalFree(out.pb_data as *mut c_void);}
        Ok(format!("dpapi:{}",b64(&enc)))
    }
}
#[cfg(not(windows))]
mod dpapi {
    pub fn protect(raw:&str)->Result<String,String>{Ok(raw.trim().to_owned())}
}
fn canonical_hotkey(raw:&str)->String{
    let t=raw.trim();
    if let Some((mods,key))=t.rsplit_once('+'){
        if key.len()==1&&key.chars().all(|c|c.is_ascii_digit()){
            let mut has_ctrl=false;let mut has_shift=false;
            for part in mods.split('+'){
                if part.eq_ignore_ascii_case("ctrl"){has_ctrl=true;}
                else if part.eq_ignore_ascii_case("shift"){has_shift=true;}
                else{return t.to_owned();}
            }
            if has_ctrl&&has_shift{return format!("Ctrl+Shift+{key}")}
            if has_ctrl{return format!("Ctrl+{key}")}
        }
    }
    t.to_owned()
}
fn is_allowed_api_url(raw:&str)->bool{
    let Ok(address)=url::Url::parse(raw.trim()) else{return false};
    if address.scheme()!="https"||address.path()!="/api/excel-daily/events"||
        !address.username().is_empty()||address.password().is_some()||address.query().is_some()||address.fragment().is_some(){
        return false
    }
    match address.host(){
        Some(url::Host::Domain(d))=>{
            let h=d.to_ascii_lowercase();
            !(h=="localhost"||h.ends_with(".localhost")||h.ends_with(".local")||h.ends_with(".internal"))
        }
        Some(url::Host::Ipv4(v4))=>{
            !(v4.is_loopback()||v4.is_private()||v4.is_link_local()||v4.is_unspecified()||v4.is_broadcast()||v4.octets()[0]==0)
        }
        _=>false
    }
}
fn dir()->Result<PathBuf,String>{
    env::current_exe().map_err(|_|"設定の保存先を取得できません。".to_owned())?
      .parent().map(Path::to_owned).ok_or_else(||"設定の保存先がありません。".to_owned())
}
fn log(root:&Path,msg:&str){
    let mut paths=vec![root.join("excel-daily-settings.log")];
    if let Some(p)=env::var_os("LOCALAPPDATA"){paths.push(PathBuf::from(p).join("ExcelDaily").join("settings.log"))}
    for p in paths{
        if let Some(parent)=p.parent(){let _=fs::create_dir_all(parent);}
        if let Ok(mut f)=fs::OpenOptions::new().create(true).append(true).open(p){let _=writeln!(f,"pid={} {}",std::process::id(),msg);}
    }
}
fn parse_cfg(raw:&str)->Result<Value,String>{
    let v:Value=serde_json::from_str(raw.trim_start_matches('\u{feff}')).map_err(|_|"settings.json の形式が正しくありません。".to_owned())?;
    if !v.is_object(){return Err("settings.json の形式が正しくありません。".into())}
    Ok(v)
}
fn read_cfg(root:&Path)->Result<Value,String>{
    let raw=fs::read_to_string(root.join("settings.json")).map_err(|_|"settings.json が見つかりません。".to_owned())?;
    parse_cfg(&raw)
}
fn text_at(v:&Value,group:&str,key:&str,default:&str)->String{
    v[group][key].as_str().unwrap_or(default).to_owned()
}
fn flag(v:&Value,group:&str,key:&str,default:bool)->bool{v[group][key].as_bool().unwrap_or(default)}
fn to_form(v:&Value)->SettingsForm{
    let defaults:Value=serde_json::from_str(DEFAULT_CFG).expect("valid default settings");
    let d=|g:&str,k:&str|text_at(&defaults,g,k,"");
    SettingsForm{
        name:text_at(v,"agent","display_name",""),url:text_at(v,"data_source","api_url",&d("data_source","api_url")),
        token:String::new(),token_configured:!text_at(v,"data_source","api_token","").trim().is_empty(),
        pattern:text_at(v,"workbook","filename_pattern",&d("workbook","filename_pattern")),
        selected:text_at(v,"workbook","selected_workbook",""),day_sheet:text_at(v,"sheets","day",&d("sheets","day")),
        night_sheet:text_at(v,"sheets","night",&d("sheets","night")),
        startup:flag(v,"app","start_with_windows",false),auto_update:flag(v,"auto_update","enabled",true),
        auto_save:flag(v,"auto_save","enabled",true),
        auto_save_minutes:Some(v["auto_save"]["interval_minutes"].as_u64().filter(|n|(1..=120).contains(n)).unwrap_or(15) as i64),
        remote_update:flag(v,"agent","remote_update",true),
        day:canonical_hotkey(&text_at(v,"hotkeys","day","Ctrl+1")),
        night:canonical_hotkey(&text_at(v,"hotkeys","night","Ctrl+2")),
        refresh:canonical_hotkey(&text_at(v,"hotkeys","refresh","Ctrl+3")),
        tomorrow:canonical_hotkey(&text_at(v,"hotkeys","tomorrow","Ctrl+4"))
    }
}
fn patch(v:&mut Value,group:&str,key:&str,value:Value)->Result<(),String>{
    let root=v.as_object_mut().ok_or("設定形式が正しくありません。")?;
    let entry=root.entry(group.to_owned()).or_insert_with(||json!({}));
    if entry.is_null(){*entry=json!({});}
    entry.as_object_mut().ok_or("設定形式が正しくありません。")?.insert(key.to_owned(),value);
    Ok(())
}
fn merge_form(mut v:Value,f:&SettingsForm)->Result<Value,String>{
    if !is_allowed_api_url(&f.url){
        return Err("Worker API URL は安全な HTTPS の /api/excel-daily/events を指定してください。".into())
    }
    let keys=hotkeys::HotkeyOptions{day:f.day.clone(),night:f.night.clone(),refresh:f.refresh.clone(),tomorrow:f.tomorrow.clone()};
    hotkeys::validate_hotkeys(&keys).map_err(|_|"ショートカットキーが重複しているか無効です。")?;
    let minutes=f.auto_save_minutes.filter(|m|(1..=120).contains(m)).ok_or("自動保存間隔は1〜120分で指定してください。")?;
    for (group,key,value) in [
        ("auto_save","enabled",json!(f.auto_save)),("auto_save","interval_minutes",json!(minutes)),
        ("agent","display_name",json!(f.name.trim())),("agent","remote_update",json!(f.remote_update)),
        ("data_source","api_url",json!(f.url.trim())),
        ("workbook","filename_pattern",json!(f.pattern)),("workbook","selected_workbook",json!(f.selected)),
        ("sheets","day",json!(f.day_sheet)),("sheets","night",json!(f.night_sheet)),
        ("app","start_with_windows",json!(f.startup)),("auto_update","enabled",json!(f.auto_update)),
        ("hotkeys","day",json!(canonical_hotkey(&f.day))),("hotkeys","night",json!(canonical_hotkey(&f.night))),
        ("hotkeys","refresh",json!(canonical_hotkey(&f.refresh))),("hotkeys","tomorrow",json!(canonical_hotkey(&f.tomorrow)))
    ]{patch(&mut v,group,key,value)?}
    if !f.token.trim().is_empty(){
        patch(&mut v,"data_source","api_token",json!(dpapi::protect(f.token.trim())?))?;
    }else{
        let existing=text_at(&v,"data_source","api_token","");
        if !existing.trim().is_empty()&&!existing.trim().starts_with("dpapi:"){
            patch(&mut v,"data_source","api_token",json!(dpapi::protect(existing.trim())?))?;
        }
    }
    Ok(v)
}
fn set_startup(root:&Path,enabled:bool)->Result<(),String>{
    let key=RegKey::predef(HKEY_CURRENT_USER).create_subkey(r"Software\Microsoft\Windows\CurrentVersion\Run")
      .map_err(|_|"Windows 自動起動の設定に失敗しました。")?.0;
    if enabled{
        key.set_value("ExcelDaily",&format!("\"{}\"",root.join("excel-daily.exe").display())).map_err(|_|"Windows 自動起動の設定に失敗しました。")?;
    }else if let Err(e)=key.delete_value("ExcelDaily"){
        if e.kind()!=std::io::ErrorKind::NotFound{return Err("Windows 自動起動の解除に失敗しました。".into())}
    }
    Ok(())
}
fn save(root:&Path,f:&SettingsForm)->Result<(),String>{
    let config=merge_form(read_cfg(root)?,f)?;
    let bytes=serde_json::to_vec_pretty(&config).map_err(|_|"設定の変換に失敗しました。")?;
    let temp=root.join(format!("settings.{}.tmp",std::process::id()));
    fs::write(&temp,bytes).map_err(|_|"設定の書き込みに失敗しました。")?;
    if fs::rename(&temp,root.join("settings.json")).is_err(){let _=fs::remove_file(temp);return Err("設定を保存できませんでした。".into())}
    set_startup(root,f.startup)?;
    Ok(())
}
fn ps()->Command{let mut c=Command::new("powershell.exe");c.args(["-NoProfile","-NonInteractive","-ExecutionPolicy","Bypass"]).creation_flags(NO_WINDOW).stdin(Stdio::null());c}
fn bridge(root:&Path,action:&str)->Result<Value,String>{
    let output=ps().arg("-Command").arg(BRIDGE).env("ED_ROOT",root).env("ED_ACTION",action).current_dir(root).output()
      .map_err(|_|"エージェント操作を開始できませんでした。")?;
    if !output.status.success(){log(root,&format!("BRIDGE_FAILED action={action}"));return Err("操作に失敗しました。更新中でないか、エージェントの状態をご確認ください。".into())}
    serde_json::from_slice(&output.stdout).map_err(|_|"エージェントの応答を取得できませんでした。".into())
}
#[tauri::command]
fn load_settings()->Result<Value,String>{let root=dir()?;Ok(json!({"version":VERSION,"form":to_form(&read_cfg(&root)?)}))}
#[tauri::command]
fn save_settings(form:SettingsForm)->Result<(),String>{save(&dir()?,&form)}
#[tauri::command]
async fn agent_action(action:String)->Result<Value,String>{
    if !["status","restart","stop","tomorrow"].contains(&action.as_str()){return Err("不明な操作です。".into())}
    tauri::async_runtime::spawn_blocking(move||bridge(&dir()?,&action)).await.map_err(|_|"操作に失敗しました。".to_owned())?
}
struct Update{path:PathBuf,child:Child}
#[derive(Default)]
struct Tracker(Mutex<Option<Update>>);
#[tauri::command]
fn start_update(tracker:tauri::State<'_,Tracker>)->Result<(),String>{
    let root=dir()?;
    let mut slot=tracker.0.lock().map_err(|_|"更新状態を取得できません。")?;
    if slot.is_some(){return Err("更新を確認中です。".into())}
    let n=SystemTime::now().duration_since(UNIX_EPOCH).map_err(|_|"時刻が正しくありません。")?.as_nanos();
    let path=root.join(format!("excel-daily-update-{}-{n}.txt",std::process::id()));
    let child=ps().arg("-File").arg(root.join("check_update.ps1")).arg("-Executable").arg(root.join("excel-daily.exe"))
      .args(["-CurrentVersion",VERSION,"-Repo","maedandori/excel-daily","-PidToWait"]).arg(std::process::id().to_string())
      .arg("-StatusFile").arg(&path).stdout(Stdio::null()).stderr(Stdio::null()).current_dir(&root).spawn()
      .map_err(|_|"更新を開始できませんでした。")?;
    *slot=Some(Update{path,child});Ok(())
}
#[tauri::command]
fn update_state(tracker:tauri::State<'_,Tracker>)->Result<String,String>{
    let mut slot=tracker.0.lock().map_err(|_|"更新状態を取得できません。")?;
    let Some(update)=slot.as_mut() else{return Ok("IDLE".into())};
    let raw=fs::read_to_string(&update.path).unwrap_or_default();
    let status=raw.trim();
    let finished=update.child.try_wait().map_err(|_|"更新状態を取得できません。")?.is_some();
    if status=="UPDATE_READY"{return Ok("UPDATE_READY".into())}
    if status=="UP_TO_DATE"||status=="UPDATE_DEFERRED"||status.starts_with("UPDATE_FAILED")||finished{
        let result=if status=="UP_TO_DATE"||status=="UPDATE_DEFERRED"{status}else{"UPDATE_FAILED"}.to_owned();
        let _=fs::remove_file(&update.path);*slot=None;return Ok(result)
    }
    Ok(if status=="DOWNLOADING"{"DOWNLOADING"}else{"CHECKING"}.into())
}
#[tauri::command]
fn close_after_update(app:tauri::AppHandle,tracker:tauri::State<'_,Tracker>)->Result<(),String>{
    let mut slot=tracker.0.lock().map_err(|_|"更新状態を取得できません。")?;
    if slot.as_ref().and_then(|u|fs::read_to_string(&u.path).ok()).is_some_and(|s|s.trim()=="UPDATE_READY"){
        if let Some(u)=slot.take(){let _=fs::remove_file(u.path);}
        app.exit(0);Ok(())
    }else{Err("更新準備が完了していません。".into())}
}
fn classic(root:&Path)->Result<(),String>{
    let script=root.join("excel-daily-settings.ps1");
    fs::write(&script,format!("\u{feff}{}",LEGACY.trim_start_matches('\u{feff}'))).map_err(|_|"旧設定画面を準備できません。")?;
    fs::write(root.join("settings_update_state.ps1"),UPDATE_STATE).map_err(|_|"旧設定画面を準備できません。")?;
    let status=ps().arg("-STA").arg("-File").arg(script).env("EXCEL_DAILY_SETTINGS_VERSION",VERSION).current_dir(root).status()
      .map_err(|_|"旧設定画面を起動できません。")?;
    if !status.success(){return Err("旧設定画面を起動できません。".into())}Ok(())
}
fn main(){
    let Ok(root)=dir() else{return};
    log(&root,&format!("TAURI_START version={VERSION}"));
    if env::args().any(|a|a=="--classic"){if classic(&root).is_err(){log(&root,"CLASSIC_FAILED")}return}
    match tauri::Builder::default().manage(Tracker::default())
        .invoke_handler(tauri::generate_handler![load_settings,save_settings,agent_action,start_update,update_state,close_after_update])
        .build(tauri::generate_context!()){
        Ok(app)=>app.run(|_,_|{}),
        Err(e)=>{log(&root,&format!("TAURI_INIT_FAILED {e}"));if classic(&root).is_err(){log(&root,"CLASSIC_FAILED")}}
    }
}
#[cfg(test)]
mod tests{
    use super::*;
    #[test]fn config_roundtrip_preserves_unknown_settings(){
        let mut v=parse_cfg(DEFAULT_CFG).unwrap();v["rows"]["step"]=json!(3);v["extra"]=json!({"value":42});
        v["data_source"]["api_token"]=json!("stored-secret");
        let mut f=to_form(&v);assert!(f.token.is_empty());assert!(f.token_configured);
        f.name="第一工場".into();
        let kept=merge_form(v.clone(),&f).unwrap();
        assert!(!kept["data_source"]["api_token"].as_str().unwrap().is_empty());
        f.token="local-only".into();
        let result=merge_form(v,&f).unwrap();assert_eq!(result["rows"]["step"],3);assert_eq!(result["extra"]["value"],42);
        assert_eq!(result["agent"]["display_name"],"第一工場");
        assert!(!result["data_source"]["api_token"].as_str().unwrap().is_empty());
    }
    #[test]fn autosave_defaults_and_roundtrip(){
        let mut v=parse_cfg(DEFAULT_CFG).unwrap();
        v.as_object_mut().unwrap().remove("auto_save");
        let mut f=to_form(&v);
        assert!(f.auto_save);assert_eq!(f.auto_save_minutes,Some(15));
        f.auto_save=false;f.auto_save_minutes=Some(30);
        let result=merge_form(v.clone(),&f).unwrap();
        assert_eq!(result["auto_save"]["enabled"],false);
        assert_eq!(result["auto_save"]["interval_minutes"],30);
        f.auto_save_minutes=Some(0);assert!(merge_form(v.clone(),&f).is_err());
        f.auto_save_minutes=None;assert!(merge_form(v,&f).is_err());
    }
    #[test]fn invalid_inputs_are_rejected(){
        let v=parse_cfg(DEFAULT_CFG).unwrap();let mut f=to_form(&v);f.tomorrow=f.day.clone();
        assert!(merge_form(v.clone(),&f).is_err());f.tomorrow="Ctrl+4".into();
        for bad in ["http://example.com/api/excel-daily/events","https://127.0.0.1/api/excel-daily/events","https://localhost/api/excel-daily/events","https://192.168.1.10/api/excel-daily/events"]{
            f.url=bad.into();assert!(merge_form(v.clone(),&f).is_err());
        }
        assert_eq!(canonical_hotkey("ctrl+shift+3"),"Ctrl+Shift+3");
    }
}
