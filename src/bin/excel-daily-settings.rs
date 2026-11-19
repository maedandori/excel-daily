#![cfg_attr(windows, windows_subsystem = "windows")]
#[cfg(not(windows))]
compile_error!("Excel Daily Settings requires Windows.");
use std::{fs,os::windows::process::CommandExt,process::{Command,Stdio}};
fn main()->Result<(),Box<dyn std::error::Error>>{
 let exe=std::env::current_exe()?;
 let dir=exe.parent().ok_or("Settings installation directory missing")?;
 let ps=dir.join("excel-daily-settings.ps1");
 fs::write(&ps,concat!("\u{feff}",include_str!("../../scripts/settings_gui.ps1")))?;
 let result=Command::new("powershell.exe")
  .args(["-NoProfile","-STA","-WindowStyle","Hidden","-ExecutionPolicy","Bypass","-File"])
  .arg(&ps)
  .env("EXCEL_DAILY_SETTINGS_VERSION",env!("CARGO_PKG_VERSION"))
  .creation_flags(0x08000000).stdin(Stdio::null()).status()?;
 if !result.success(){return Err(format!("Settings UI failed: {result}").into());}
 Ok(())
}
