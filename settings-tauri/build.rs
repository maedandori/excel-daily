use std::{fs, path::PathBuf};
#[path="../branding/app_icon.rs"]
mod app_icon;
fn main() {
    println!("cargo:rerun-if-changed=../branding/app_icon.rs");
    let root = PathBuf::from(std::env::var("CARGO_MANIFEST_DIR").unwrap());
    let icons = root.join("icons");
    fs::create_dir_all(&icons).unwrap();
    app_icon::write_icon(&icons.join("icon.ico"));
    tauri_build::build();
}
