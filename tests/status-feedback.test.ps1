$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$agent=Get-Content -LiteralPath (Join-Path $root 'src\main.rs') -Raw -Encoding UTF8
$ribbon=Get-Content -LiteralPath (Join-Path $root 'ribbon\RibbonController.cs') -Raw -Encoding UTF8
$ui=Get-Content -LiteralPath (Join-Path $root 'ribbon\StatusUi.cs') -Raw -Encoding UTF8
$script=Get-Content -LiteralPath (Join-Path $root 'scripts\update_excel.ps1') -Raw -Encoding UTF8
$status=Get-Content -LiteralPath (Join-Path $root 'src\status.rs') -Raw -Encoding UTF8
if($agent -match 'MessageBoxW|confirm_tomorrow|\binfo\('){throw 'Agent still has a modal operation message'}
if($ribbon -match 'MessageBox.Show'){throw 'Ribbon must not show modal dialogs'}
if($script -match '\bInputBox\b'){throw 'Excel refresh must not show modal workbook picker'}
if($ribbon -notmatch "getEnabled='IsCommandEnabled'"){throw 'Ribbon getEnabled callback absent'}
if(($ribbon.Split("getEnabled='IsCommandEnabled'").Count-1) -lt 5){throw 'Every Ribbon button must disable while busy'}
if($ribbon -notmatch 'StatusUi.Guard\('){throw 'Copied worksheet Shape must reject duplicate clicks'}
if($ui -notmatch 'excel-daily-busy.lock' -or $ui -notmatch 'excel-daily-command.json'){throw 'UI is not guarding queue and busy lock'}
if($ui -notmatch 'Invalidate\('){throw 'Ribbon does not refresh disabled appearance'}
if($ui -notmatch '\.StatusBar\s*='){throw 'Excel StatusBar is not used'}
if($status -notmatch 'excel-daily-status.json'){throw 'Agent feedback channel is absent'}
foreach($key in @('"running"','"success"','"error"')){
 if($agent -notmatch [regex]::Escape($key)){throw "Agent status transition $key is missing"}
}
Write-Output 'STATUS_FEEDBACK_TESTS_OK'

$autosave=Get-Content -LiteralPath (Join-Path $root 'ribbon\WorkbookAutoSave.cs') -Raw -Encoding UTF8
if($autosave -notmatch 'excel-daily-recovery\.lock' -or $autosave -notmatch 'excel-daily-busy\.lock'){throw 'Auto save must respect recovery and busy markers'}
if($script -notmatch 'excel-daily-recovery\.lock'){throw 'Persistent recovery marker missing'}
if($script.IndexOf('[IO.File]::WriteAllText($recoveryLock') -gt $script.IndexOf('foreach ($e in $ordered)')){throw 'Recovery marker must precede workbook changes'}
if($script.IndexOf('Remove-Item -LiteralPath $recoveryLock') -lt $script.IndexOf('$book.Save()')){throw 'Recovery marker removed too early'}
Write-Output 'RECOVERY_GUARD_TESTS_OK'
