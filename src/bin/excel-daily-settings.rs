#![cfg_attr(windows, windows_subsystem = "windows")]
#[cfg(not(windows))]
compile_error!("Excel Daily Settings requires Windows.");
use std::{fs,io::Write,os::windows::process::CommandExt,path::PathBuf,process::{Command,Stdio}};
use windows_sys::Win32::UI::WindowsAndMessaging::{MessageBoxW,MB_OK,MB_ICONERROR};
const CREATE_NO_WINDOW:u32=0x08000000;
fn wide(s:&str)->Vec<u16>{s.encode_utf16().chain(std::iter::once(0)).collect()}
fn report_problem(message:&str) {
 let title=wide("Excel Daily Settings Error");
 let body=wide(message);
 unsafe{MessageBoxW(std::ptr::null_mut(),body.as_ptr(),title.as_ptr(),MB_OK|MB_ICONERROR);}
}
fn run() -> Result<(),String> {
 let exe=std::env::current_exe().map_err(|e|e.to_string())?;
 let dir=exe.parent().ok_or("Settings installation directory missing")?;
 let script=dir.join("excel-daily-settings.ps1");
 // A BOM is essential to preserve Japanese text in Windows PowerShell 5.1.
 fs::write(&script,concat!("\u{feff}",include_str!("../../scripts/settings_gui.ps1")))
  .map_err(|e|format!("Cannot create Settings script: {e}"))?;
 let result=Command::new("powershell.exe")
  .args(["-NoProfile","-STA","-ExecutionPolicy","Bypass","-File"])
  .arg(&script)
  .env("EXCEL_DAILY_SETTINGS_VERSION",env!("CARGO_PKG_VERSION"))
  .creation_flags(CREATE_NO_WINDOW)
  .stdin(Stdio::null())
  .output().map_err(|e|format!("Cannot start PowerShell: {e}"))?;
 if !result.status.success() {
  let err=String::from_utf8_lossy(&result.stderr);
  let out=String::from_utf8_lossy(&result.stdout);
  let details=if !err.trim().is_empty(){err.to_string()}else{out.to_string()};
  let short:String=details.chars().take(1300).collect();
  return Err(format!("Settings exited: {}\n{}\nSee excel-daily-settings.log.",result.status,short));
 }
 Ok(())
}
fn main(){
 if let Err(e)=run(){
  let dir=std::env::current_exe().ok()
   .and_then(|p|p.parent().map(PathBuf::from))
   .unwrap_or_else(||PathBuf::from("."));
  if let Ok(mut file)=fs::OpenOptions::new().append(true).create(true)
   .open(dir.join("excel-daily-settings.log")){
    let _=writeln!(file,"ERROR: {e}");
  }
  report_problem(&e);
  std::process::exit(1);
 }
}
