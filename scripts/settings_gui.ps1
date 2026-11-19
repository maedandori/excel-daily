param([switch]$SelfTest,[string]$ConfigPath='')
$ErrorActionPreference='Stop'
$dir=if($ConfigPath){Split-Path -Parent ([IO.Path]::GetFullPath($ConfigPath))}else{$PSScriptRoot}
$file=if($ConfigPath){[IO.Path]::GetFullPath($ConfigPath)}else{Join-Path $dir 'settings.json'}
$log=Join-Path $dir 'excel-daily-settings.log'
try {
 Add-Type -AssemblyName System.Windows.Forms
 Add-Type -AssemblyName System.Drawing
 [System.Windows.Forms.Application]::EnableVisualStyles()
 if(!(Test-Path -LiteralPath $file)){throw "Missing settings.json: $file"}
 $script:Cfg=Get-Content -LiteralPath $file -Raw -Encoding UTF8|ConvertFrom-Json
 if(!$script:Cfg.agent){$script:Cfg|Add-Member -NotePropertyName agent -NotePropertyValue ([pscustomobject]@{display_name='';remote_update=$true})}
 if(!$script:Cfg.hotkeys){$script:Cfg|Add-Member -NotePropertyName hotkeys -NotePropertyValue ([pscustomobject]@{day='Ctrl+1';night='Ctrl+2';refresh='Ctrl+3';tomorrow='Ctrl+4'})}
 foreach($entry in @(@('day','Ctrl+1'),@('night','Ctrl+2'),@('refresh','Ctrl+3'),@('tomorrow','Ctrl+4'))){
  $key=$entry[0]
  if(!$script:Cfg.hotkeys.PSObject.Properties[$key]){$script:Cfg.hotkeys|Add-Member -NotePropertyName $key -NotePropertyValue $entry[1]}
 }
 $version=[string]$env:EXCEL_DAILY_SETTINGS_VERSION
 if(!$version){$version='unknown'}
 $agentVersion='not started'
 $versionFile=Join-Path $dir 'excel-daily-version.json'
 if(Test-Path $versionFile){
  try{
   $a=Get-Content $versionFile -Raw -Encoding UTF8|ConvertFrom-Json
   if($a.version){$agentVersion="v$($a.version) (last started)"}
   $proc=Get-Process -Id ([int]$a.pid) -ErrorAction SilentlyContinue
   if($proc -and $proc.ProcessName -ieq 'excel-daily'){$agentVersion="v$($a.version) (running)"}
  }catch{}
 }
 $form=[System.Windows.Forms.Form]::new()
 $form.Text="Excel Daily Settings v$version / 設定"
 $form.ClientSize=[System.Drawing.Size]::new(700,760)
 $form.StartPosition=[System.Windows.Forms.FormStartPosition]::CenterScreen
 $form.FormBorderStyle=[System.Windows.Forms.FormBorderStyle]::FixedSingle
 $form.MaximizeBox=$false;$form.ShowInTaskbar=$true
 $script:Fields=@{};$script:HotkeyFields=@{}
 function LabelAt([string]$t,[int]$x,[int]$y,[int]$w=180){
  $v=[System.Windows.Forms.Label]::new();$v.Text=$t
  $v.Location=[System.Drawing.Point]::new($x,$y)
  $v.Size=[System.Drawing.Size]::new($w,28)
  $form.Controls.Add($v);return $v
 }
 function TextAt([string]$t,[string]$k,[string]$v,[int]$y,[bool]$secret=$false){
  $null=LabelAt $t 20 $y 175
  $box=[System.Windows.Forms.TextBox]::new()
  $box.Location=[System.Drawing.Point]::new(202,$y)
  $box.Size=[System.Drawing.Size]::new(471,25)
  $box.Text=$v;$box.UseSystemPasswordChar=$secret
  $form.Controls.Add($box);$script:Fields[$k]=$box
 }
 $header=LabelAt ("Settings v"+$version+"     |     Agent "+$agentVersion) 20 14 665
 $header.Font=[System.Drawing.Font]::new('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
 TextAt 'Device name / 端末名' 'name' ([string]$script:Cfg.agent.display_name) 55
 TextAt 'Worker API URL' 'url' ([string]$script:Cfg.data_source.api_url) 90
 TextAt 'Excel Daily token' 'token' ([string]$script:Cfg.data_source.api_token) 125 $true
 TextAt 'Workbook pattern' 'pattern' ([string]$script:Cfg.workbook.filename_pattern) 160
 TextAt 'Selected workbook' 'selected' ([string]$script:Cfg.workbook.selected_workbook) 195
 TextAt 'Day sheet / 昼勤' 'daySheet' ([string]$script:Cfg.sheets.day) 230
 TextAt 'Night sheet / 夜勤' 'nightSheet' ([string]$script:Cfg.sheets.night) 265
 function CheckAt([string]$t,[int]$x,[bool]$value){
  $check=[System.Windows.Forms.CheckBox]::new()
  $check.Text=$t;$check.Location=[System.Drawing.Point]::new($x,304)
  $check.Size=[System.Drawing.Size]::new(220,27);$check.Checked=$value
  $form.Controls.Add($check);return $check
 }
 $script:Startup=CheckAt 'Start with Windows' 20 ([bool]$script:Cfg.app.start_with_windows)
 $script:Auto=CheckAt 'Auto GitHub update' 249 ([bool]$script:Cfg.auto_update.enabled)
 $script:Remote=CheckAt 'Admin update requests' 478 ($script:Cfg.agent.remote_update -ne $false)
 $null=LabelAt 'HOTKEYS - phím tắt không được trùng nhau' 20 340 650
 $choices=[System.Collections.Generic.List[string]]::new()
 foreach($i in 0..9){$choices.Add("Ctrl+$i")}
 foreach($i in 0..9){$choices.Add("Ctrl+Shift+$i")}
 function HotkeyAt([string]$label,[string]$key,[int]$x,[int]$y){
  $null=LabelAt $label $x $y 145
  $combo=[System.Windows.Forms.ComboBox]::new()
  $combo.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
  $combo.Location=[System.Drawing.Point]::new(($x+145),$y)
  $combo.Size=[System.Drawing.Size]::new(164,27)
  $combo.Items.AddRange([object[]]$choices.ToArray())
  $combo.SelectedItem=[string]$script:Cfg.hotkeys.$key
  if($combo.SelectedIndex -lt 0){throw "Invalid configured hotkey for $key"}
  $form.Controls.Add($combo);$script:HotkeyFields[$key]=$combo
 }
 HotkeyAt 'Ca ngày / 昼勤' 'day' 20 379
 HotkeyAt 'Ca đêm / 夜勤' 'night' 366 379
 HotkeyAt 'Chỉ cập nhật giờ' 'refresh' 20 420
 HotkeyAt 'Ngày mai / 翌日' 'tomorrow' 366 420
 $null=LabelAt 'Ngày mai: lấy lịch cả ca ngày và đêm của NGÀY MAI, không xóa các dòng đã có.' 20 470 665
 $null=LabelAt 'Mở đúng workbook dành cho ngày mai trước khi nhập.' 20 497 665
 $script:Status=LabelAt 'Ready. Save and restart Agent to apply new shortcuts.' 20 709 665
 function SaveSettings {
  if($script:Fields['url'].Text -notmatch '^https://[^/]+/api/excel-daily/events$'){throw 'Invalid Worker API URL'}
  $used=@()
  foreach($key in @('day','night','refresh','tomorrow')){
   $value=[string]$script:HotkeyFields[$key].SelectedItem
   if(!$value -or $used -contains $value){throw "Hotkeys must be unique and selected: $key"}
   $used+=$value;$script:Cfg.hotkeys.$key=$value
  }
  $script:Cfg.agent.display_name=$script:Fields['name'].Text.Trim()
  $script:Cfg.agent.remote_update=[bool]$script:Remote.Checked
  $script:Cfg.data_source.api_url=$script:Fields['url'].Text.Trim()
  $script:Cfg.data_source.api_token=$script:Fields['token'].Text.Trim()
  $script:Cfg.workbook.filename_pattern=$script:Fields['pattern'].Text
  $script:Cfg.workbook.selected_workbook=$script:Fields['selected'].Text
  $script:Cfg.sheets.day=$script:Fields['daySheet'].Text
  $script:Cfg.sheets.night=$script:Fields['nightSheet'].Text
  $script:Cfg.app.start_with_windows=[bool]$script:Startup.Checked
  $script:Cfg.auto_update.enabled=[bool]$script:Auto.Checked
  [IO.File]::WriteAllText($file,(ConvertTo-Json -InputObject $script:Cfg -Depth 25),([Text.UTF8Encoding]::new($false)))
 }
 function StopOwnAgent {
  $agentPath=Join-Path $dir 'excel-daily.exe'
  Get-CimInstance Win32_Process -Filter "name='excel-daily.exe'" -ErrorAction SilentlyContinue |
   Where-Object { $_.ExecutablePath -ieq $agentPath } |
   ForEach-Object {Stop-Process -Id $_.ProcessId -ErrorAction SilentlyContinue}
 }
 function ButtonAt([string]$t,[int]$x,[int]$y,[int]$w,[scriptblock]$cb){
  $b=[System.Windows.Forms.Button]::new()
  $b.Text=$t;$b.Location=[System.Drawing.Point]::new($x,$y)
  $b.Size=[System.Drawing.Size]::new($w,36)
  $b.Add_Click($cb);$form.Controls.Add($b)
 }
 ButtonAt 'Save / 保存' 20 542 150 {
  try{SaveSettings;$script:Status.Text='Saved. Restart Agent to activate shortcuts.'}
  catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null}
 }
 ButtonAt 'Restart Agent' 186 542 150 {
  try {
   SaveSettings;StopOwnAgent;Start-Sleep -Milliseconds 700
   Start-Process -FilePath (Join-Path $dir 'excel-daily.exe') -WorkingDirectory $dir
   $script:Status.Text='Agent restarted with saved shortcuts.'
  }catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null}
 }
 ButtonAt 'Stop Agent' 352 542 150 {
  StopOwnAgent;$script:Status.Text='Agent stopped.'
 }
 ButtonAt 'Close' 518 542 160 {$form.Close()}
 ButtonAt 'Ghi lịch ngày mai / 翌日' 20 594 310 {
  try{
   if([System.Windows.Forms.MessageBox]::Show(
    'Use the workbook for TOMORROW. Do not use today''s workbook. Continue?',
    'Excel Daily - Tomorrow',
    [System.Windows.Forms.MessageBoxButtons]::YesNo
   ) -ne [System.Windows.Forms.DialogResult]::Yes){return}
   $livePath=Join-Path $dir 'excel-daily-version.json'
   if(!(Test-Path $livePath)){throw 'Start the updated Agent first.'}
   $live=Get-Content $livePath -Raw -Encoding UTF8|ConvertFrom-Json
   if([version]([string]$live.version) -lt [version]'0.3.2'){throw 'Agent v0.3.2+ required.'}
   $proc=Get-Process -Id ([int]$live.pid) -ErrorAction SilentlyContinue
   if(!$proc -or $proc.ProcessName -ine 'excel-daily'){throw 'Agent is not running.'}
   $commandPath=Join-Path $dir 'excel-daily-command.json'
   if(Test-Path $commandPath){throw 'A previous request is still pending.'}
   [IO.File]::WriteAllText($commandPath,'{"action":"tomorrow"}',([Text.UTF8Encoding]::new($false)))
   $script:Status.Text='Tomorrow import sent. Confirm the Agent popup.'
  }catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null}
 }
 ButtonAt 'GitHub Update / 更新' 346 594 332 {
  try{
   SaveSettings
   $helper=Join-Path $dir 'check_update.ps1'
   $statusPath=Join-Path $dir 'excel-daily-update-status.txt'
   if(Test-Path $statusPath){Remove-Item $statusPath -Force}
   $args=@('-NoProfile','-NonInteractive','-WindowStyle','Hidden','-ExecutionPolicy','Bypass','-File',('"' + $helper + '"'),
    '-Executable',('"' + (Join-Path $dir 'excel-daily.exe') + '"'),
    '-CurrentVersion',$version,'-Repo','maedandori/excel-daily',
    '-PidToWait',[string]$PID,'-StatusFile',('"' + $statusPath + '"'))
   Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList $args|Out-Null
   $script:Status.Text='Checking GitHub Release...'
  }catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null}
 }
 $timer=[System.Windows.Forms.Timer]::new();$timer.Interval=1000
 $timer.Add_Tick({
  $p=Join-Path $dir 'excel-daily-update-status.txt'
  if(Test-Path $p){
   try{$message=([IO.File]::ReadAllText($p)).Trim()
    if($message){$script:Status.Text=$message}
    if($message -eq 'UPDATE_READY'){$timer.Stop();$form.Close()}
   }catch{}
  }
 })
 if($SelfTest){
  if($script:HotkeyFields.Count -ne 4 -or $form.Controls.Count -lt 15){throw 'Settings UI did not initialize completely'}
  Write-Output "SETTINGS_SELFTEST_OK version=$version controls=$($form.Controls.Count)"
  $form.Dispose()
  return
 }
 $timer.Start()
 $form.Add_Shown({$form.Activate();$form.BringToFront()})
 [void]$form.ShowDialog()
}catch{
 $details="Startup error: "+$_.Exception.Message+" (line "+$_.InvocationInfo.ScriptLineNumber+")"
 try{Add-Content -LiteralPath $log -Value ("$(Get-Date -Format o) "+$details) -Encoding UTF8}catch{}
 try{[System.Windows.Forms.MessageBox]::Show($details+' See excel-daily-settings.log.','Excel Daily Settings Error')|Out-Null}catch{}
 Write-Error $details
 exit 1
}
