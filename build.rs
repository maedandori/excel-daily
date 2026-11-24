use std::{env, path::PathBuf};
#[path="branding/app_icon.rs"]
mod app_icon;
fn main() {
    println!("cargo:rerun-if-changed=branding/app_icon.rs");
    if env::var("CARGO_CFG_TARGET_OS").as_deref() != Ok("windows") { return; }
    let out = PathBuf::from(env::var("OUT_DIR").unwrap());
    let path = out.join("excel-daily.ico");
    app_icon::write_icon(&path);
    let mut res = winres::WindowsResource::new();
    res.set_icon(path.to_str().unwrap());
    let version = env::var("CARGO_PKG_VERSION").unwrap_or_else(|_|"0.0.0".to_owned());
    res.set("FileVersion", &version);
    res.set("ProductVersion", &version);
    res.set("ProductName", "Excel Daily");
    res.set("FileDescription", "MICS Excel Daily Hotkey Agent");
    res.compile().unwrap();
}
