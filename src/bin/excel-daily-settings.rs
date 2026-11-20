#![cfg_attr(windows, windows_subsystem = "windows")]
#[cfg(not(windows))]
compile_error!("Excel Daily Settings requires Windows.");
use std::{env,fs,io::Write,os::windows::process::CommandExt,path::{Path,PathBuf},process::{Command,Stdio}};
#[path = "../status.rs"]
mod status;
const CREATE_NO_WINDOW:u32=0x08000000;
fn log_event(dir:&Path,message:&str){
 let mut paths=vec![dir.join("excel-daily-settings.log")];
 if let Some(local)=env::var_os("LOCALAPPDATA") {
  paths.push(PathBuf::from(local).join("ExcelDaily").join("settings.log"));
 }
 paths.push(env::temp_dir().join("excel-daily-settings.log"));
 for p in paths {
  if let Some(parent)=p.parent(){let _=fs::create_dir_all(parent);}
  if let Ok(mut file)=fs::OpenOptions::new().append(true).create(true).open(&p){
   let _=writeln!(file,"pid={} {}",std::process::id(),message);
  }
 }
}
fn settings_script_contents()->String {
 let source=include_str!("../../scripts/settings_gui.ps1");
 format!("\u{feff}{}",source.trim_start_matches('\u{feff}'))
}
#[cfg(test)]
mod tests{
 use super::*;
 #[test] fn embedded_settings_has_exactly_one_bom(){
  let contents=settings_script_contents();
  assert!(contents.starts_with('\u{feff}'));
  assert_ne!(contents.chars().nth(1),Some('\u{feff}'));
  assert!(contents.trim_start_matches('\u{feff}').starts_with("param("));
 }
}
fn run(dir:&Path) -> Result<(),String> {
 let script=dir.join("excel-daily-settings.ps1");
 // The source may already start with a BOM. Exactly one BOM must be written:
 // a duplicated BOM can make Windows PowerShell 5.1 reject the param() declaration.
 fs::write(&script,settings_script_contents())
  .map_err(|e|format!("Cannot create Settings script: {e}"))?;
 fs::write(dir.join("settings_update_state.ps1"),include_str!("../../scripts/settings_update_state.ps1"))
  .map_err(|e|format!("Cannot create update-state script: {e}"))?;
 log_event(dir,"POWERSHELL_START");
 let result=Command::new("powershell.exe")
  .args(["-NoProfile","-STA","-ExecutionPolicy","Bypass","-File"])
  .arg(&script)
  .env("EXCEL_DAILY_SETTINGS_VERSION",env!("CARGO_PKG_VERSION"))
  .creation_flags(CREATE_NO_WINDOW)
  .stdin(Stdio::null())
  .output().map_err(|e|format!("Cannot start PowerShell: {e}"))?;
 log_event(dir,&format!("POWERSHELL_EXIT status={}",result.status));
 if !result.status.success() {
  let err=String::from_utf8_lossy(&result.stderr);
  let out=String::from_utf8_lossy(&result.stdout);
  let details=if !err.trim().is_empty(){err.to_string()}else{out.to_string()};
  let short:String=details.chars().take(1300).collect();
  return Err(format!("Settings exited: {}\n{}\nSee settings.log in LOCALAPPDATA\\ExcelDaily.",result.status,short));
 }
 Ok(())
}
fn main(){
 let dir=env::current_exe().ok().and_then(|p|p.parent().map(PathBuf::from))
   .unwrap_or_else(||PathBuf::from("."));
 log_event(&dir,&format!("EXE_START version={}",env!("CARGO_PKG_VERSION")));
 std::panic::set_hook(Box::new(|p|{
  let dir=env::current_exe().ok().and_then(|f|f.parent().map(PathBuf::from)).unwrap_or_default();
  log_event(&dir,&format!("RUST_PANIC {}",p));
 }));
 match run(&dir){
  Ok(())=>log_event(&dir,"EXE_EXIT_NORMAL"),
  Err(e)=>{
   log_event(&dir,&format!("EXE_EXIT_ERROR {e}"));
   status::write(&dir,"error","設定画面を起動できませんでした。詳細はexcel-daily-settings.logを確認してください。");
   std::process::exit(1);
  }
 }
}
