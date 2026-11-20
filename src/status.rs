use std::{fs, path::Path, time::{SystemTime, UNIX_EPOCH}};

pub fn status_message(text: &str) -> String {
    text.replace('\r', " ").replace('\n', " ").chars().take(190).collect::<String>().trim().to_owned()
}

pub fn write(dir: &Path, state: &str, message: &str) {
    let now = SystemTime::now().duration_since(UNIX_EPOCH)
        .unwrap_or_default().as_millis();
    let content = serde_json::json!({
        "state": state,
        "message": status_message(message),
        "updated_ms": now
    });
    let temp = dir.join("excel-daily-status.tmp");
    let path = dir.join("excel-daily-status.json");
    if let Ok(bytes) = serde_json::to_vec(&content) {
        if fs::write(&temp, bytes).is_ok() {
            // Windows rename doesn't overwrite a destination. If Excel's UI
            // reads while this file is briefly absent it will retry next tick.
            let _ = fs::remove_file(&path);
            let _ = fs::rename(&temp, &path);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn status_has_no_newlines_or_excessively_long_text() {
        let m = status_message(&format!("a\rb\nc{}", "x".repeat(400)));
        assert!(m.starts_with("a b c"));
        assert!(m.chars().count() <= 190);
        assert!(!m.contains('\n'));
    }
    #[test]
    fn status_is_utf8_json_with_result_state() {
        let dir = std::env::temp_dir().join(format!("excel-daily-status-test-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        write(&dir, "success", "更新完了 — hoàn tất");
        let v: serde_json::Value = serde_json::from_slice(
            &fs::read(dir.join("excel-daily-status.json")).unwrap()).unwrap();
        assert_eq!(v["state"], "success");
        assert!(v["message"].as_str().unwrap().contains("更新完了"));
        assert!(v["updated_ms"].as_u64().unwrap() > 0);
        fs::remove_dir_all(&dir).unwrap();
    }
}
