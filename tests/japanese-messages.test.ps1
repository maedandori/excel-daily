$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$paths=@('ribbon/StatusUi.cs','ribbon/RibbonController.cs','ribbon/SheetButtonContract.cs',
 'src/main.rs','scripts/update_excel.ps1','scripts/settings_gui.ps1',
 'scripts/install-ribbon.ps1','scripts/apply_update.ps1')
$messages=($paths|ForEach-Object{Get-Content -LiteralPath (Join-Path $root $_) -Raw -Encoding UTF8}) -join [Environment]::NewLine
foreach($old in @('đang cập nhật','không sửa Excel','Lưu file Excel','Hãy kiểm tra','hoàn tất',
   'Hãy mở file','Agent restarted','Tomorrow import sent','Confirm the Agent popup',
   'Không tìm thấy','Chỉ cập nhật giờ','Cannot register hotkey')){
 if($messages.Contains($old)){throw "Old user-visible Vietnamese/English message: $old"}
}
$settings=Get-Content -LiteralPath (Join-Path $root 'scripts/settings_gui.ps1') -Raw -Encoding UTF8
if($settings -match 'MessageBox'){throw 'Settings must not show modal operation messages'}
$launcher=Get-Content -LiteralPath (Join-Path $root 'src/bin/excel-daily-settings.rs') -Raw -Encoding UTF8
if($launcher -match 'MessageBox'){throw 'Settings launcher must not use modal dialogs'}
$ribbon=Get-Content -LiteralPath (Join-Path $root 'ribbon/StatusUi.cs') -Raw -Encoding UTF8
if($ribbon -notmatch '更新中'){throw 'Japanese Ribbon feedback missing'}
$agent=Get-Content -LiteralPath (Join-Path $root 'src/main.rs') -Raw -Encoding UTF8
if($agent -notmatch 'UpdateDecision::Deferred' -or $agent -notmatch 'Duration::from_secs\(if deferred'){throw 'Deferred periodic retry missing'}
Write-Output 'JAPANESE_MESSAGES_TESTS_OK'
