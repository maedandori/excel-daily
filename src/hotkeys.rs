use serde::Deserialize;

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct HotkeyOptions {
    pub day: String,
    pub night: String,
    pub refresh: String,
    pub tomorrow: String,
}
impl Default for HotkeyOptions {
    fn default() -> Self {
        Self {
            day: "Ctrl+1".to_owned(),
            night: "Ctrl+2".to_owned(),
            refresh: "Ctrl+3".to_owned(),
            tomorrow: "Ctrl+4".to_owned(),
        }
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Shortcut {
    pub modifiers: u32,
    pub key: u32,
}
/// Allow Ctrl+digit or Ctrl+Shift+digit, matching Settings dropdown.
pub fn parse_hotkey(value: &str) -> Result<Shortcut, String> {
    let chunks: Vec<_> = value.split('+').map(|p| p.trim().to_ascii_uppercase()).collect();
    let (modifiers, key) = match chunks.as_slice() {
        [ctrl, digit] if ctrl == "CTRL" => (0x0002, digit),
        [ctrl, shift, digit] if ctrl == "CTRL" && shift == "SHIFT" => (0x0002 | 0x0004, digit),
        _ => return Err(format!("Invalid hotkey: {value}. Use Ctrl+0..9 or Ctrl+Shift+0..9.")),
    };
    if key.len() != 1 || !key.as_bytes()[0].is_ascii_digit() {
        return Err(format!("Invalid hotkey key: {value}"));
    }
    Ok(Shortcut { modifiers, key: key.as_bytes()[0] as u32 })
}
pub fn validate_hotkeys(config: &HotkeyOptions) -> Result<[Shortcut; 4], String> {
    let values = [&config.day, &config.night, &config.refresh, &config.tomorrow];
    let mut codes = Vec::with_capacity(4);
    for value in values { codes.push(parse_hotkey(value)?); }
    for i in 0..codes.len() {
        if codes[i + 1..].contains(&codes[i]) {
            return Err(format!("Duplicate hotkey: {}", values[i]));
        }
    }
    Ok([codes[0], codes[1], codes[2], codes[3]])
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test] fn defaults_are_distinct() {
        assert_eq!(validate_hotkeys(&HotkeyOptions::default()).unwrap().len(), 4);
    }
    #[test] fn accepts_and_rejects_combinations() {
        assert_eq!(parse_hotkey("Ctrl+Shift+9").unwrap(), Shortcut{modifiers:6,key:57});
        assert!(parse_hotkey("Alt+4").is_err());
        assert!(parse_hotkey("Ctrl+F5").is_err());
        let mut config = HotkeyOptions::default();
        config.tomorrow = config.day.clone();
        assert!(validate_hotkeys(&config).is_err());
    }
}
