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
use windows_sys::Win32::UI::WindowsAndMessaging::{GetMessageW, MSG, WM_HOTKEY};
use windows_sys::Win32::UI::Input::KeyboardAndMouse::{RegisterHotKey, UnregisterHotKey};
use winreg::{enums::HKEY_CURRENT_USER, RegKey};
mod hotkeys;
mod status;
use hotkeys::{HotkeyOptions,validate_hotkeys};

const DAY_ID: i32 = 0x5101;
const NIGHT_ID: i32 = 0x5102;
const REFRESH_ID: i32 = 0x5103;
const TOMORROW_ID: i32 = 0x5104;

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
    agent: AgentOptions,
    hotkeys: HotkeyOptions,
}
#[derive(Debug, Deserialize)]
#[serde(default)]
struct AgentOptions { remote_update: bool }
impl Default for AgentOptions { fn default()->Self{Self{remote_update:true}} }
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
const CREATE_NO_WINDOW: u32 = 0x08000000;
// Operation commands never block Excel with Windows message boxes. The
// XLL reads progress, success and errors from excel-daily-status.json.
fn call_excel(dir: PathBuf, action: &'static str, busy: Arc<AtomicBool>, workbook_target: Option<String>, sheet_target: Option<String>) {
    let script = dir.join("update_excel.ps1");
    let config = dir.join("settings.json");
    let busy_file = dir.join("excel-daily-busy.lock");
    let _ = fs::write(&busy_file, std::process::id().to_string());
    status::write(&dir, "running", "更新中：生産予定を取得し、Excelに反映しています。");
    // Ribbon/Shape commands are scoped to ONE clicked worksheet.
    // Legacy Ctrl+1/2/3/4 hotkeys retain their original shift behavior.
    let shifts: &[&str] = if sheet_target.is_some() {
        &["auto"]
    } else if matches!(action, "all" | "refresh_times" | "tomorrow") {
        &["day", "night"]
    } else { &[action] };
    let mut failures: Vec<String> = Vec::new();
    let mut results: Vec<String> = Vec::new();
    for (index, shift) in shifts.iter().enumerate() {
        status::write(&dir, "running", &format!(
            "更新中：{}/{}（{}）", index + 1, shifts.len(),
            if *shift == "auto" { "選択中のシート" } else if *shift == "day" { "昼勤" } else { "夜勤" }));
        let mut command = Command::new("powershell.exe");
        command.args(["-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden",
            "-ExecutionPolicy", "Bypass", "-File"])
            .arg(&script).arg("-Settings").arg(&config).arg("-Shift").arg(shift)
            .creation_flags(CREATE_NO_WINDOW).stdin(Stdio::null());
        if action == "refresh_times" { command.arg("-TimeOnly"); }
        if action == "tomorrow" { command.arg("-DateOffset").arg("1"); }
        if let Some(ref full_name) = workbook_target {
            command.arg("-WorkbookFullName").arg(full_name);
        }
        if let Some(ref sheet_name) = sheet_target {
            command.arg("-WorksheetName").arg(sheet_name);
        }
        match command.output() {
            Ok(o) => {
                let out = String::from_utf8_lossy(&o.stdout);
                let err = String::from_utf8_lossy(&o.stderr);
                log(&dir, &format!("{} action={} status={} stdout={} stderr={}",
                    shift, action, o.status, out.trim(), err.trim()));
                if o.status.success() {
                    let summary = out.lines().find(|line|line.contains("added=")).unwrap_or("");
                    let count = |name: &str| -> String {
                        summary.split(';').find_map(|part| {
                            part.trim().strip_prefix(name).and_then(|x|
                                x.trim().split_whitespace().next())
                                .filter(|x| x.chars().all(|c| c.is_ascii_digit()))
                                .map(str::to_owned)
                        }).unwrap_or_else(|| "0".to_owned())
                    };
                    let label = if *shift == "auto" {
                        sheet_target.as_deref().unwrap_or("選択中のシート")
                    } else if *shift == "day" { "昼勤" } else { "夜勤" };
                    results.push(format!("{label}：新規{}件・時刻変更{}件・保持{}件",
                        count("added="), count("updated_times="), count("preserved=")));
                } else {
                    failures.push(format!("{}の更新に失敗しました",
                        if *shift == "auto" { "選択中のシート" } else if *shift == "day" { "昼勤" } else { "夜勤" }));
                }
            }
            Err(e) => {
                log(&dir, &format!("{} action={} powershell failed: {}", shift, action, e));
                failures.push(format!("{}の更新に失敗しました",
                    if *shift == "day" { "昼勤" } else { "夜勤" }));
            }
        }
    }
    if failures.is_empty() {
        status::write(&dir, "success",
            &format!("更新完了：{}。", results.join("／")));
    } else {
        status::write(&dir, "error",
            &format!("更新エラー：{}。詳細はexcel-daily.logを確認してください。", failures.join("、")));
    }
    let _ = fs::remove_file(&busy_file);
    busy.store(false, Ordering::SeqCst);
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum UpdateDecision { Ready, Deferred, NoChange }

fn check_for_update(dir: &Path, exe: &Path, repo: &str) -> UpdateDecision {
    let mut command = Command::new("powershell.exe");
    let output = command
        .args(["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File"])
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
                status::write(dir, "error", "更新の確認に失敗しました。詳細はexcel-daily.logを確認してください。");
            } else if stdout.contains("UPDATE_DEFERRED") {
                log(dir, "UPDATE_DEFERRED: Excel open or workbook task active.");
                status::write(dir, "deferred", "自動更新を保留しました。Excelを閉じると再確認します。");
                return UpdateDecision::Deferred;
            } else if stdout.contains("UPDATE_READY") {
                log(dir, "A verified update is staged; exiting for replacement.");
                status::write(dir, "running", "更新ファイルの確認が完了しました。安全に更新を適用しています。");
                return UpdateDecision::Ready;
            }
        }
        Err(e) => {
            log(dir, &format!("Update check unavailable: {e}"));
            status::write(dir, "error", "更新の確認を開始できませんでした。詳細はログを確認してください。");
        }
    }
    UpdateDecision::NoChange
}
// Ribbon is installed per-user, without administrative permissions or VBA.
// Do not treat Office Add-in policy failures as fatal for the original hotkeys.
fn ensure_ribbon_registered(dir: &Path) {
    let script = dir.join("install-ribbon.ps1");
    if !script.is_file() {
        log(dir, "RIBBON_INSTALL_MISSING_SCRIPT: use complete excel-daily-windows.zip.");
        return;
    }
    if !dir.join("ExcelDailyRibbon.xll").exists() || !dir.join("ExcelDailyRibbon64.xll").exists() {
        log(dir, "RIBBON_INSTALL_MISSING_XLL: install complete ZIP, not the standalone EXE.");
        return;
    }
    let result = Command::new("powershell.exe")
        .args(["-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass", "-File"])
        .arg(&script)
        .creation_flags(CREATE_NO_WINDOW)
        .stdin(Stdio::null())
        .output();
    match result {
        Ok(output) => {
            let stdout = String::from_utf8_lossy(&output.stdout);
            let stderr = String::from_utf8_lossy(&output.stderr);
            if output.status.success() {
                log(dir, &format!("RIBBON_REGISTER_OK: {}", stdout.trim()));
            } else {
                log(dir, &format!("RIBBON_REGISTER_FAILED: code={} stdout={} stderr={}",
                    output.status, stdout.trim(), stderr.trim()));
            }
        }
        Err(e) => log(dir, &format!("RIBBON_INSTALL_LAUNCH_FAILED: {e}")),
    }
}

fn agent_presence(dir: &Path,command_id: Option<u64>,status: Option<&str>)->Option<serde_json::Value>{
    let mut p=Command::new("powershell.exe");
    p.args(["-NoProfile","-WindowStyle","Hidden","-ExecutionPolicy","Bypass","-File"])
     .arg(dir.join("agent_presence.ps1")).arg("-Settings").arg(dir.join("settings.json"))
     .arg("-Version").arg(env!("CARGO_PKG_VERSION"))
     .creation_flags(CREATE_NO_WINDOW).stdin(Stdio::null());
    if let (Some(id),Some(status))=(command_id,status) {
       p.arg("-CommandId").arg(id.to_string()).arg("-CommandStatus").arg(status);
    }
    match p.output(){
       Ok(out) if out.status.success()=>serde_json::from_slice(&out.stdout).ok(),
       Ok(out)=>{log(dir,&format!("Presence failed: {}",String::from_utf8_lossy(&out.stderr).trim()));None},
       Err(e)=>{log(dir,&format!("Presence invocation failed: {e}"));None}
    }
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
    fs::write(dir.join("round_time.ps1"), include_str!("../scripts/round_time.ps1"))?;
    fs::write(dir.join("shift_date.ps1"), include_str!("../scripts/shift_date.ps1"))?;
    fs::write(dir.join("color_cells.ps1"), include_str!("../scripts/color_cells.ps1"))?;
    fs::write(dir.join("agent_presence.ps1"), include_str!("../scripts/agent_presence.ps1"))?;
    fs::write(dir.join("check_update.ps1"), include_str!("../scripts/check_update.ps1"))?;
    fs::write(dir.join("apply_update.ps1"), include_str!("../scripts/apply_update.ps1"))?;
    fs::write(dir.join("update_safety.ps1"), include_str!("../scripts/update_safety.ps1"))?;
    fs::write(dir.join("install-ribbon.ps1"), include_str!("../scripts/install-ribbon.ps1"))?;
    ensure_ribbon_registered(&dir);
    let version_path=dir.join("excel-daily-version.json");
    let version_info=serde_json::json!({"version":env!("CARGO_PKG_VERSION"),"pid":std::process::id()});
    let _=fs::write(&version_path,serde_json::to_string_pretty(&version_info)?);
    log(&dir, &format!("Agent started. v{}",env!("CARGO_PKG_VERSION")));
    if !cfg.app.run_in_background || cfg.app.show_status_on_start {
        log(&dir, "Agent ready; Excel Ribbon shows status without modal startup dialogs.");
    }
    let shortcuts=match validate_hotkeys(&cfg.hotkeys){
        Ok(codes)=>codes,
        Err(e)=>{log(&dir,&format!("Hotkey configuration error: {e}"));status::write(&dir,"error","ショートカットキーの設定が正しくありません。設定画面をご確認ください。");anyhow::bail!(e)}
    };
    let configured=[
        (DAY_ID,cfg.hotkeys.day.as_str(),shortcuts[0]),
        (NIGHT_ID,cfg.hotkeys.night.as_str(),shortcuts[1]),
        (REFRESH_ID,cfg.hotkeys.refresh.as_str(),shortcuts[2]),
        (TOMORROW_ID,cfg.hotkeys.tomorrow.as_str(),shortcuts[3])
    ];
    let mut registered=Vec::new();
    for (id,label,shortcut) in configured {
        let ok=unsafe{RegisterHotKey(std::ptr::null_mut(),id,shortcut.modifiers,shortcut.key)};
        if ok==0 {
            for prev in registered {unsafe{UnregisterHotKey(std::ptr::null_mut(),prev)};}
            let msg=format!("ショートカットキー「{}」を登録できません。他のアプリと重複していないか設定画面で確認してください。",label);
            log(&dir,&msg);status::write(&dir,"error",&msg);
            anyhow::bail!(msg);
        }
        registered.push(id);
    }
    let busy = Arc::new(AtomicBool::new(false));
    if cfg.auto_update.enabled {
        let update_busy = busy.clone();
        let update_dir = dir.clone();
        let update_exe = exe.clone();
        let update_repo = cfg.auto_update.repository.clone();
        let hours = cfg.auto_update.check_interval_hours.clamp(1, 168);
        thread::spawn(move || {
            // Excel running: retry cheaply (no network) until user closes it.
            thread::sleep(Duration::from_secs(60));
            loop {
                let mut deferred = false;
                if !update_busy.swap(true, Ordering::SeqCst) {
                    match check_for_update(&update_dir, &update_exe, &update_repo) {
                        UpdateDecision::Ready => std::process::exit(0),
                        UpdateDecision::Deferred => deferred = true,
                        UpdateDecision::NoChange => {}
                    }
                    update_busy.store(false, Ordering::SeqCst);
                }
                thread::sleep(Duration::from_secs(if deferred { 120 } else { hours * 3600 }));
            }
        });
    }
    if cfg.agent.remote_update {
        let dir2=dir.clone();
        let exe2=exe.clone();
        let repo=cfg.auto_update.repository.clone();
        let busy2=busy.clone();
        thread::spawn(move || {
            thread::sleep(Duration::from_secs(15));
            loop {
                if !busy2.swap(true,Ordering::SeqCst) {
                    if let Some(res)=agent_presence(&dir2,None,None){
                        if let Some(cmd)=res["update_command_id"].as_u64().filter(|v|*v>0){
                            log(&dir2,&format!("Admin requested update id={cmd}"));
                            match check_for_update(&dir2,&exe2,&repo) {
                                UpdateDecision::Ready => {
                                    let _=agent_presence(&dir2,Some(cmd),Some("staged"));
                                    std::process::exit(0);
                                }
                                UpdateDecision::NoChange => {
                                    let _=agent_presence(&dir2,Some(cmd),Some("failed"));
                                }
                                UpdateDecision::Deferred => {
                                    // Keep remote command pending; retry when Excel closes.
                                    log(&dir2,"Admin update deferred; command left pending.");
                                }
                            }
                        }
                    }
                    busy2.store(false,Ordering::SeqCst);
                }
                thread::sleep(Duration::from_secs(90));
            }
        });
    }
    // Settings UI requests are processed by this Agent under the same busy lock
    // as global hotkeys. Never edit the workbook concurrently.
    {
        let queue_dir=dir.clone();
        let queue_busy=busy.clone();
        thread::spawn(move||loop {
            let p=queue_dir.join("excel-daily-command.json");
            if p.is_file() && !queue_busy.swap(true,Ordering::SeqCst) {
                let action_value=fs::read_to_string(&p).ok().and_then(|s|serde_json::from_str::<serde_json::Value>(&s).ok());
                let _=fs::remove_file(&p);
                let requested=action_value.as_ref().and_then(|v|v["action"].as_str());
                let allowed=match requested {
                    Some("all")=>Some("all"),
                    Some("day")=>Some("day"),
                    Some("night")=>Some("night"),
                    Some("refresh_times")=>Some("refresh_times"),
                    Some("tomorrow")=>Some("tomorrow"),
                    _=>None,
                };
                if let Some(action)=allowed {
                    log(&queue_dir,&format!("Ribbon/Settings requested action={action}"));
                    // A Shape carries its clicked workbook path, not a fixed
                    // machine-specific XLL path. Never fall back to another
                    // workbook when the target can't be found.
                    let workbook_target=action_value.as_ref()
                        .and_then(|v|v["workbook_full_name"].as_str())
                        .map(str::to_owned);
                    let sheet_target=action_value.as_ref()
                        .and_then(|v|v["sheet_name"].as_str())
                        .map(str::to_owned);
                    if let Some(ref full_name)=workbook_target {
                        if !std::path::Path::new(full_name).is_absolute() {
                            queue_busy.store(false,Ordering::SeqCst);
                            log(&queue_dir,"Invalid Ribbon workbook target; request dropped.");
                            status::write(&queue_dir,"error","指定されたExcelファイルが無効です。更新は実行されませんでした。");
                            continue;
                        }
                    }
                    // Old XLLs may send a workbook with no sheet. Never
                    // silently fall back to both sheets after an upgrade.
                    if (sheet_target.is_some() && workbook_target.is_none()) ||
                        (workbook_target.is_some() &&
                         matches!(action,"all"|"refresh_times") && sheet_target.is_none()) ||
                        sheet_target.as_ref().is_some_and(|name|
                            name.is_empty() || name.encode_utf16().count()>31) {
                        queue_busy.store(false,Ordering::SeqCst);
                        log(&queue_dir,"Workbook/sheet routing missing or invalid; request rejected.");
                        status::write(&queue_dir,"error","シートの指定が無効です。Excelを再起動してから再実行してください。");
                        continue;
                    }
                    call_excel(queue_dir.clone(),action,queue_busy.clone(),workbook_target,sheet_target);
                } else {
                    queue_busy.store(false,Ordering::SeqCst);
                    log(&queue_dir,"Invalid Ribbon/Settings command ignored.");
                    status::write(&queue_dir,"error","更新コマンドが無効です。ファイルは変更されていません。");
                }
            }
            thread::sleep(Duration::from_secs(2));
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
                REFRESH_ID => "refresh_times",
                TOMORROW_ID => "tomorrow",
                _ => continue,
            };
            if !busy.swap(true, Ordering::SeqCst) {
                let dir2 = dir.clone();
                let busy2 = busy.clone();
                thread::spawn(move || call_excel(dir2, shift, busy2, None, None));
            } else {
                log(&dir, "Ignored duplicate hotkey while update active.");
            }
        }
    }
    unsafe {
        UnregisterHotKey(std::ptr::null_mut(), DAY_ID);
        UnregisterHotKey(std::ptr::null_mut(), NIGHT_ID);
        UnregisterHotKey(std::ptr::null_mut(), REFRESH_ID);
        UnregisterHotKey(std::ptr::null_mut(), TOMORROW_ID);
    }
    Ok(())
}
