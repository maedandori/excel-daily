#![cfg_attr(windows, windows_subsystem = "windows")]
#[cfg(not(windows))]
compile_error!("excel-daily requires Windows and Microsoft Excel.");

use std::{
    fs::{self, OpenOptions},
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::atomic::{AtomicBool, Ordering},
    sync::Arc,
    thread,
    time::Duration,
    os::windows::process::CommandExt,
};
use serde::Deserialize;
use windows_sys::Win32::UI::WindowsAndMessaging::{
    GetMessageW, MessageBoxW, MSG, MB_ICONINFORMATION, MB_OK,
    WM_HOTKEY,
};
use windows_sys::Win32::UI::Input::KeyboardAndMouse::{RegisterHotKey, UnregisterHotKey};
const MOD_CONTROL: u32 = 0x0002;
use winreg::{enums::HKEY_CURRENT_USER, RegKey};

const DAY_ID: i32 = 0x5101;
const NIGHT_ID: i32 = 0x5102;

#[derive(Debug, Default, Deserialize)]
#[serde(default)]
struct AppOptions {
    run_in_background: bool,
    start_with_windows: bool,
    show_status_on_start: bool,
}
#[derive(Debug, Default, Deserialize)]
#[serde(default)]
struct Config {
    app: AppOptions,
    auto_update: UpdateOptions,
}
#[derive(Debug, Deserialize)]
#[serde(default)]
struct UpdateOptions {
    enabled: bool,
    check_interval_hours: u64,
    repository: String,
}
impl Default for UpdateOptions {
    fn default() -> Self { Self { enabled: false, check_interval_hours: 6, repository: "maedandori/excel-daily".to_string() } }
}
fn executable_dir() -> PathBuf {
    std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(Path::to_path_buf))
        .unwrap_or_else(|| PathBuf::from("."))
}
fn log(dir: &Path, message: &str) {
    let path = dir.join("excel-daily.log");
    if let Ok(mut f) = OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(f, "{:?}: {}", std::time::SystemTime::now(), message);
    }
}
fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}
fn info(msg: &str) {
    let msg = wide(msg);
    let title = wide("Excel Daily");
    unsafe { MessageBoxW(std::ptr::null_mut(), msg.as_ptr(), title.as_ptr(), MB_OK | MB_ICONINFORMATION); }
}
fn set_autostart(enabled: bool, exe: &Path) -> anyhow::Result<()> {
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);
    let (run, _) = hkcu.create_subkey("Software\\Microsoft\\Windows\\CurrentVersion\\Run")?;
    if enabled {
        run.set_value("ExcelDaily", &format!("\"{}\"", exe.display()))?;
    } else {
        match run.delete_value("ExcelDaily") {
            Ok(()) => (),
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => (),
            Err(e) => return Err(e.into()),
        }
    }
    Ok(())
}
fn call_excel(dir: PathBuf, shift: &'static str, busy: Arc<AtomicBool>) {
    let script = dir.join("update_excel.ps1");
    let config = dir.join("settings.json");
    let output = Command::new("powershell.exe")
        .args(["-NoProfile", "-ExecutionPolicy", "Bypass", "-File"])
        .arg(&script)
        .args(["-Settings"])
        .arg(&config)
        .args(["-Shift", shift])
        .stdin(Stdio::null())
        .output();
    match output {
        Ok(o) => {
            let out = String::from_utf8_lossy(&o.stdout);
            let err = String::from_utf8_lossy(&o.stderr);
            log(&dir, &format!("{} status={} stdout={} stderr={}", shift, o.status, out.trim(), err.trim()));
            if !o.status.success() { info(&format!("Excel Daily lỗi khi cập nhật {}.\nXem excel-daily.log.", shift)); }
        }
        Err(e) => {
            log(&dir, &format!("powershell failed: {e}"));
            info("Không chạy được PowerShell. Xem excel-daily.log.");
        }
    }
    busy.store(false, Ordering::SeqCst);
}
const CREATE_NO_WINDOW: u32 = 0x08000000;
fn check_for_update(dir: &Path, exe: &Path, repo: &str) -> bool {
    let mut command = Command::new("powershell.exe");
    let output = command
        .args(["-NoProfile", "-ExecutionPolicy", "Bypass", "-File"])
        .arg(dir.join("check_update.ps1"))
        .arg("-Executable").arg(exe)
        .arg("-CurrentVersion").arg(env!("CARGO_PKG_VERSION"))
        .arg("-Repo").arg(repo)
        .arg("-PidToWait").arg(std::process::id().to_string())
        .creation_flags(CREATE_NO_WINDOW)
        .stdin(Stdio::null())
        .output();
    match output {
        Ok(o) => {
            let stdout = String::from_utf8_lossy(&o.stdout);
            if !o.status.success() {
                log(dir, &format!("Update check failed: {}", String::from_utf8_lossy(&o.stderr)));
            } else if stdout.contains("UPDATE_READY") {
                log(dir, "A verified update is staged; exiting for replacement.");
                return true;
            }
        }
        Err(e) => log(dir, &format!("Update check unavailable: {e}")),
    }
    false
}
fn main() -> anyhow::Result<()> {
    let dir = executable_dir();
    let cfg_path = dir.join("settings.json");
    let cfg: Config = serde_json::from_str(&fs::read_to_string(&cfg_path)
        .map_err(|e| anyhow::anyhow!("Cannot read {}: {}", cfg_path.display(), e))?)?;
    let exe = std::env::current_exe()?;
    if let Err(e) = set_autostart(cfg.app.start_with_windows, &exe) {
        log(&dir, &format!("autostart registry error: {e}"));
    }
    // Embed scripts so the ZIP cannot accidentally ship mismatched helpers.
    fs::write(dir.join("update_excel.ps1"), include_str!("../scripts/update_excel.ps1"))?;
    fs::write(dir.join("check_update.ps1"), include_str!("../scripts/check_update.ps1"))?;
    fs::write(dir.join("apply_update.ps1"), include_str!("../scripts/apply_update.ps1"))?;
    log(&dir, "Agent started.");
    if !cfg.app.run_in_background || cfg.app.show_status_on_start {
        info("Excel Daily đang chạy.\nCtrl+1: ca ngày\nCtrl+2: ca đêm\nThay đổi thiết lập trong settings.json, sau đó khởi động lại ứng dụng.");
    }
    let day_ok = unsafe { RegisterHotKey(std::ptr::null_mut(), DAY_ID, MOD_CONTROL as u32, '1' as u32) };
    let night_ok = unsafe { RegisterHotKey(std::ptr::null_mut(), NIGHT_ID, MOD_CONTROL as u32, '2' as u32) };
    if day_ok == 0 || night_ok == 0 {
        log(&dir, "RegisterHotKey failed: hotkey already in use.");
        if day_ok != 0 { unsafe { UnregisterHotKey(std::ptr::null_mut(), DAY_ID); } }
        if night_ok != 0 { unsafe { UnregisterHotKey(std::ptr::null_mut(), NIGHT_ID); } }
        info("Ctrl+1 hoặc Ctrl+2 đã được ứng dụng khác sử dụng.");
        anyhow::bail!("Global hotkey conflict");
    }
    let busy = Arc::new(AtomicBool::new(false));
    if cfg.auto_update.enabled {
        let update_busy = busy.clone();
        let update_dir = dir.clone();
        let update_exe = exe.clone();
        let update_repo = cfg.auto_update.repository.clone();
        let hours = cfg.auto_update.check_interval_hours.clamp(1, 168);
        thread::spawn(move || {
            // Wait until startup is complete. Updates never interrupt Excel writes.
            thread::sleep(Duration::from_secs(60));
            loop {
                if !update_busy.swap(true, Ordering::SeqCst) {
                    if check_for_update(&update_dir, &update_exe, &update_repo) {
                        // The detached helper waits for this process to exit.
                        std::process::exit(0);
                    }
                    update_busy.store(false, Ordering::SeqCst);
                }
                thread::sleep(Duration::from_secs(hours * 3600));
            }
        });
    }
    let mut msg: MSG = unsafe { std::mem::zeroed() };
    loop {
        let rc = unsafe { GetMessageW(&mut msg, std::ptr::null_mut(), 0, 0) };
        if rc <= 0 { break; }
        if msg.message == WM_HOTKEY {
            let shift = match msg.wParam as i32 {
                DAY_ID => "day",
                NIGHT_ID => "night",
                _ => continue,
            };
            if !busy.swap(true, Ordering::SeqCst) {
                let dir2 = dir.clone();
                let busy2 = busy.clone();
                thread::spawn(move || call_excel(dir2, shift, busy2));
            } else {
                log(&dir, "Ignored duplicate hotkey while update active.");
            }
        }
    }
    unsafe {
        UnregisterHotKey(std::ptr::null_mut(), DAY_ID);
        UnregisterHotKey(std::ptr::null_mut(), NIGHT_ID);
    }
    Ok(())
}
