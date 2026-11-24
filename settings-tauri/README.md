# Tauri Settings

Windows-only Settings interface. Build from repo root using
`cargo build --manifest-path settings-tauri/Cargo.toml --release --target x86_64-pc-windows-msvc`.
The frontend is embedded; no local HTTP server or npm build is required. WebView2 Runtime is needed.
The portable release replaces only the existing settings executable with this one; the Agent and Ribbon use their original contracts. If Tauri initialization fails, the existing WinForms Settings opens instead. To request the classic interface explicitly, run `excel-daily-settings.exe --classic`.
Configuration uses the existing settings.json and preserves unknown fields. No token is sent to a remote service by the UI. Save, startup, restart, stop, tomorrow and verified GitHub update use the existing behavior; no modal notifications. CI checks backend logic and builds the Windows executable, but a real desktop launch and Excel test remain necessary.
