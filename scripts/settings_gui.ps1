# Local Excel Daily Settings UI. Password token is never logged or sent to admin.
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$dir=$PSScriptRoot
$file=Join-Path $dir 'settings.json'
$cfg=Get-Content -LiteralPath $file -Raw -Encoding UTF8|ConvertFrom-Json
if(!$cfg.agent){$cfg|Add-Member -NotePropertyName agent -NotePropertyValue ([pscustomobject]@{display_name='';remote_update=$true})}
$form=New-Object Windows.Forms.Form
$form.Text='Excel Daily Settings / 設定'
$form.Size=New-Object Drawing.Size(650,690)
$form.FormBorderStyle='FixedDialog';$form.MaximizeBox=$false;$form.StartPosition='CenterScreen'
$fields=@{}
function Field([string]$label,[string]$key,[string]$value,[int]$y,[bool]$masked=$false){
 $lab=New-Object Windows.Forms.Label
 $lab.Text=$label;$lab.Location=New-Object Drawing.Point(18,$y)
 $lab.Size=New-Object Drawing.Size(190,24)
 $form.Controls.Add($lab)
 $box=New-Object Windows.Forms.TextBox
 $box.Location=New-Object Drawing.Point(205,$y)
 $box.Size=New-Object Drawing.Size(415,24)
 $box.Text=$value;$box.UseSystemPasswordChar=$masked
 $form.Controls.Add($box);$fields[$key]=$box
}
Field 'Device name / 端末名' 'name' ([string]$cfg.agent.display_name) 20
Field 'Worker events API URL' 'url' ([string]$cfg.data_source.api_url) 57
Field 'Excel Daily token' 'token' ([string]$cfg.data_source.api_token) 94 $true
Field 'Workbook pattern' 'book' ([string]$cfg.workbook.filename_pattern) 131
Field 'Selected workbook' 'selected' ([string]$cfg.workbook.selected_workbook) 168
Field 'Day sheet / 昼勤' 'day' ([string]$cfg.sheets.day) 205
Field 'Night sheet / 夜勤' 'night' ([string]$cfg.sheets.night) 242
$startup=New-Object Windows.Forms.CheckBox
$startup.Text='Start with Windows';$startup.Location=New-Object Drawing.Point(20,287);$startup.Size=New-Object Drawing.Size(280,27)
$startup.Checked=[bool]$cfg.app.start_with_windows;$form.Controls.Add($startup)
$auto=New-Object Windows.Forms.CheckBox
$auto.Text='Auto GitHub updates';$auto.Location=New-Object Drawing.Point(325,287);$auto.Size=New-Object Drawing.Size(270,27)
$auto.Checked=[bool]$cfg.auto_update.enabled;$form.Controls.Add($auto)
$remote=New-Object Windows.Forms.CheckBox
$remote.Text='Receive update requests from Machine-Control admin'
$remote.Location=New-Object Drawing.Point(20,321);$remote.Size=New-Object Drawing.Size(560,27)
$remote.Checked=($cfg.agent.remote_update -ne $false);$form.Controls.Add($remote)
$help=New-Object Windows.Forms.Label
$help.Location=New-Object Drawing.Point(20,362);$help.Size=New-Object Drawing.Size(590,65)
$help.Text="Ctrl+1 day, Ctrl+2 night, Ctrl+3 update pending times on both sheets.`r`nTimes are rounded UP to 10 min. Excel must be open.`r`nExisting Excel rows and app state are preserved."
$form.Controls.Add($help)
$status=New-Object Windows.Forms.Label
$status.Location=New-Object Drawing.Point(20,444);$status.Size=New-Object Drawing.Size(590,54)
$status.Text='Ready.';$form.Controls.Add($status)
function Btn([string]$label,[int]$x,[int]$y,[scriptblock]$callback){
 $button=New-Object Windows.Forms.Button
 $button.Text=$label;$button.Location=New-Object Drawing.Point($x,$y)
 $button.Size=New-Object Drawing.Size(190,36);$button.Add_Click($callback);$form.Controls.Add($button)
}
function SaveSettings {
 if($fields['url'].Text -notmatch '^https://[^/]+/api/excel-daily/events$'){throw 'Invalid Worker events URL'}
 $cfg.agent.display_name=$fields['name'].Text.Trim()
 $cfg.agent.remote_update=[bool]$remote.Checked
 $cfg.data_source.api_url=$fields['url'].Text.Trim()
 $cfg.data_source.api_token=$fields['token'].Text.Trim()
 $cfg.workbook.filename_pattern=$fields['book'].Text
 $cfg.workbook.selected_workbook=$fields['selected'].Text
 $cfg.sheets.day=$fields['day'].Text;$cfg.sheets.night=$fields['night'].Text
 $cfg.app.start_with_windows=[bool]$startup.Checked
 $cfg.auto_update.enabled=[bool]$auto.Checked
 [IO.File]::WriteAllText($file,(ConvertTo-Json -InputObject $cfg -Depth 25),(New-Object Text.UTF8Encoding($false)))
}
function StopOwnAgent {
 $agent=Join-Path $dir 'excel-daily.exe'
 Get-CimInstance Win32_Process -Filter "name='excel-daily.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.ExecutablePath -ieq $agent } |
   ForEach-Object {Stop-Process -Id $_.ProcessId -ErrorAction SilentlyContinue}
}
Btn 'Save / 保存' 18 509 {
 try{SaveSettings;$status.Text='Saved. Restart Agent to apply changes.'}
 catch{[Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null}
}
Btn 'Restart Agent' 225 509 {
 try{SaveSettings;StopOwnAgent;Start-Sleep -Milliseconds 700;Start-Process -FilePath (Join-Path $dir 'excel-daily.exe') -WorkingDirectory $dir
  $status.Text='Saved and Agent restarted.'}
 catch{[Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null}
}
Btn 'Stop Agent' 432 509 {StopOwnAgent;$status.Text='Agent stopped.'}
Btn 'GitHub Update / 更新' 18 558 {
 try{
  SaveSettings
  $script=Join-Path $dir 'check_update.ps1'
  $statusPath=Join-Path $dir 'excel-daily-update-status.txt'
  if(Test-Path $statusPath){Remove-Item -LiteralPath $statusPath -Force}
  $args=@('-NoProfile','-NonInteractive','-WindowStyle','Hidden','-ExecutionPolicy','Bypass','-File',('"' + $script + '"'),
   '-Executable',('"' + (Join-Path $dir 'excel-daily.exe') + '"'),
   '-CurrentVersion',$env:EXCEL_DAILY_SETTINGS_VERSION,
   '-Repo','maedandori/excel-daily','-PidToWait',[string]$PID,
   '-StatusFile',('"' + $statusPath + '"'))
  Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList $args|Out-Null
  $status.Text='Checking GitHub Release...'
 }catch{[Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null}
}
Btn 'Close' 432 558 {$form.Close()}
$timer=New-Object Windows.Forms.Timer;$timer.Interval=1000
$timer.Add_Tick({
 $p=Join-Path $dir 'excel-daily-update-status.txt'
 if(Test-Path $p){
  $s=([IO.File]::ReadAllText($p)).Trim();if($s){$status.Text=$s}
  if($s -eq 'UPDATE_READY'){$timer.Stop();$form.Close()}
 }
})
$timer.Start()
[void]$form.ShowDialog()
