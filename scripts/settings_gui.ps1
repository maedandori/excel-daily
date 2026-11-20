param([switch]$SelfTest,[string]$ConfigPath='')
$ErrorActionPreference='Stop'
$dir=if($ConfigPath){Split-Path -Parent ([IO.Path]::GetFullPath($ConfigPath))}else{$PSScriptRoot}
$file=if($ConfigPath){[IO.Path]::GetFullPath($ConfigPath)}else{Join-Path $dir 'settings.json'}
$log=Join-Path $dir 'excel-daily-settings.log'
. (Join-Path $PSScriptRoot 'settings_update_state.ps1')
$script:UpdateRequestActive=$false
$script:UpdateStatusPath=''
$script:ShutdownReason='normal-window-close'
function WriteSettingsLog([string]$event,[string]$detail=''){
 $stamp=Get-Date -Format o
 $entry="$stamp pid=$PID $event $detail"
 $paths=@($log)
 if($env:LOCALAPPDATA){$paths+=Join-Path $env:LOCALAPPDATA 'ExcelDaily\settings.log'}
 if($env:TEMP){$paths+=Join-Path $env:TEMP 'excel-daily-settings.log'}
 foreach($target in $paths){
  try{
   $parent=Split-Path -Parent $target
   if(!(Test-Path -LiteralPath $parent)){New-Item -ItemType Directory -Path $parent -Force|Out-Null}
   Add-Content -LiteralPath $target -Value $entry -Encoding UTF8
  }catch{}
 }
}
WriteSettingsLog 'SCRIPT_START'
try {
 Add-Type -AssemblyName System.Windows.Forms
 Add-Type -AssemblyName System.Drawing
 [System.Windows.Forms.Application]::EnableVisualStyles()
 if(!(Test-Path -LiteralPath $file)){throw "設定ファイルが見つかりません: $file"}
 $script:Cfg=Get-Content -LiteralPath $file -Raw -Encoding UTF8|ConvertFrom-Json
 if(!$script:Cfg.agent){$script:Cfg|Add-Member -NotePropertyName agent -NotePropertyValue ([pscustomobject]@{display_name='';remote_update=$true})}
 if(!$script:Cfg.hotkeys){$script:Cfg|Add-Member -NotePropertyName hotkeys -NotePropertyValue ([pscustomobject]@{day='Ctrl+1';night='Ctrl+2';refresh='Ctrl+3';tomorrow='Ctrl+4'})}
 foreach($entry in @(@('day','Ctrl+1'),@('night','Ctrl+2'),@('refresh','Ctrl+3'),@('tomorrow','Ctrl+4'))){
  $key=$entry[0]
  if(!$script:Cfg.hotkeys.PSObject.Properties[$key]){$script:Cfg.hotkeys|Add-Member -NotePropertyName $key -NotePropertyValue $entry[1]}
 }
 $version=[string]$env:EXCEL_DAILY_SETTINGS_VERSION
 if(!$version){$version='不明'}
 $agentVersion='未起動'
 $versionFile=Join-Path $dir 'excel-daily-version.json'
 if(Test-Path $versionFile){
  try{
   $a=Get-Content $versionFile -Raw -Encoding UTF8|ConvertFrom-Json
   if($a.version){$agentVersion="v$($a.version)（前回起動）"}
   $proc=Get-Process -Id ([int]$a.pid) -ErrorAction SilentlyContinue
   if($proc -and $proc.ProcessName -ieq 'excel-daily'){$agentVersion="v$($a.version)（起動中）"}
  }catch{}
 }
 $form=[System.Windows.Forms.Form]::new()
 $form.Text="Excel Daily 設定 v$version"
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
 $header=LabelAt ("設定 v"+$version+"     |     エージェント "+$agentVersion) 20 14 665
 $header.Font=[System.Drawing.Font]::new('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
 TextAt '端末名' 'name' ([string]$script:Cfg.agent.display_name) 55
 TextAt 'Worker API URL' 'url' ([string]$script:Cfg.data_source.api_url) 90
 TextAt 'Excel Daily 認証キー' 'token' ([string]$script:Cfg.data_source.api_token) 125 $true
 TextAt 'ファイル名のパターン' 'pattern' ([string]$script:Cfg.workbook.filename_pattern) 160
 TextAt '更新対象ファイル' 'selected' ([string]$script:Cfg.workbook.selected_workbook) 195
 TextAt '昼勤シート' 'daySheet' ([string]$script:Cfg.sheets.day) 230
 TextAt '夜勤シート' 'nightSheet' ([string]$script:Cfg.sheets.night) 265
 function CheckAt([string]$t,[int]$x,[bool]$value){
  $check=[System.Windows.Forms.CheckBox]::new()
  $check.Text=$t;$check.Location=[System.Drawing.Point]::new($x,304)
  $check.Size=[System.Drawing.Size]::new(220,27);$check.Checked=$value
  $form.Controls.Add($check);return $check
 }
 $script:Startup=CheckAt 'Windows起動時' 20 ([bool]$script:Cfg.app.start_with_windows)
 $script:Auto=CheckAt '自動更新を有効' 249 ([bool]$script:Cfg.auto_update.enabled)
 $script:Remote=CheckAt '管理者による更新' 478 ($script:Cfg.agent.remote_update -ne $false)
 $null=LabelAt 'ショートカットキー（重複不可）' 20 340 650
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
  if($combo.SelectedIndex -lt 0){throw "ショートカットキーの設定が正しくありません: $key"}
  $form.Controls.Add($combo);$script:HotkeyFields[$key]=$combo
 }
 HotkeyAt '昼勤の予定' 'day' 20 379
 HotkeyAt '夜勤の予定' 'night' 366 379
 HotkeyAt '時刻のみ更新' 'refresh' 20 420
 HotkeyAt '翌日の予定' 'tomorrow' 366 420
 $null=LabelAt '翌日の予定：昼勤・夜勤の両方を取得し、既存データは削除しません。' 20 470 665
 $null=LabelAt '実行前に翌日用のExcelファイルを開いてください。' 20 497 665
 $script:Status=LabelAt '準備完了：保存後にエージェントを再起動してください。' 20 709 665
 function SaveSettings {
  if($script:Fields['url'].Text -notmatch '^https://[^/]+/api/excel-daily/events$'){throw 'Worker APIのURLが正しくありません。'}
  $used=@()
  foreach($key in @('day','night','refresh','tomorrow')){
   $value=[string]$script:HotkeyFields[$key].SelectedItem
   if(!$value -or $used -contains $value){throw "ショートカットキーが重複しているか未選択です: $key"}
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
  if(Test-Path -LiteralPath (Join-Path $dir 'excel-daily-busy.lock')){throw '更新中のため停止できません。'}
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
 ButtonAt '設定を保存' 20 542 150 {
  try{SaveSettings;$script:Status.Text='設定を保存しました。再起動すると適用されます。'}
  catch{$script:Status.Text='エラー: '+$_.Exception.Message;WriteSettingsLog 'GUI_ERROR' $_.Exception.Message}
 }
 ButtonAt '再起動' 186 542 150 {
  try {
   SaveSettings;StopOwnAgent;Start-Sleep -Milliseconds 700
   Start-Process -FilePath (Join-Path $dir 'excel-daily.exe') -WorkingDirectory $dir
   $script:Status.Text='エージェントを再起動しました。'
  }catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null}
 }
 ButtonAt '停止' 352 542 150 {
  try{StopOwnAgent;$script:Status.Text='エージェントを停止しました。'}
  catch{$script:Status.Text='エラー: '+$_.Exception.Message;WriteSettingsLog 'STOP_FAILED' $_.Exception.Message}
 }
 ButtonAt '閉じる' 518 542 160 {$script:ShutdownReason='close-button';$form.Close()}
 ButtonAt '翌日の予定を取り込む' 20 594 310 {
  try{
   # 翌日の予定更新は明示的なボタン操作で実行します。
   $livePath=Join-Path $dir 'excel-daily-version.json'
   if(!(Test-Path $livePath)){throw '先にExcel Dailyエージェントを起動してください。'}
   $live=Get-Content $livePath -Raw -Encoding UTF8|ConvertFrom-Json
   if([version]([string]$live.version) -lt [version]'0.3.2'){throw 'Excel Daily v0.3.2以降が必要です。'}
   $proc=Get-Process -Id ([int]$live.pid) -ErrorAction SilentlyContinue
   if(!$proc -or $proc.ProcessName -ine 'excel-daily'){throw 'エージェントが起動していません。'}
   $commandPath=Join-Path $dir 'excel-daily-command.json'
   if(Test-Path $commandPath){throw '前の更新処理が完了していません。'}
   [IO.File]::WriteAllText($commandPath,'{"action":"tomorrow"}',([Text.UTF8Encoding]::new($false)))
   $script:Status.Text='翌日の予定更新を受け付けました。進行状況はExcel下部に表示されます。'
  }catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null}
 }
 ButtonAt 'GitHub更新を確認' 346 594 332 {
  try{
   SaveSettings
   $helper=Join-Path $dir 'check_update.ps1'
   # Unique per request: a stale UPDATE_READY from yesterday cannot close this window.
   $statusPath=Join-Path $dir ("excel-daily-update-"+[guid]::NewGuid().ToString('N')+".txt")
   $script:UpdateStatusPath=$statusPath
   $script:UpdateRequestActive=$true
   $args=@('-NoProfile','-NonInteractive','-WindowStyle','Hidden','-ExecutionPolicy','Bypass','-File',('"' + $helper + '"'),
    '-Executable',('"' + (Join-Path $dir 'excel-daily.exe') + '"'),
    '-CurrentVersion',$version,'-Repo','maedandori/excel-daily',
    '-PidToWait',[string]$PID,'-StatusFile',('"' + $statusPath + '"'))
   Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList $args|Out-Null
   $script:Status.Text='GitHubで更新を確認しています…'
   WriteSettingsLog 'UPDATE_REQUESTED' $script:UpdateStatusPath
  }catch{
   $script:UpdateRequestActive=$false
   WriteSettingsLog 'UPDATE_START_FAILED' $_.Exception.Message
   [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Excel Daily Settings')|Out-Null
 }
 }
 $timer=[System.Windows.Forms.Timer]::new();$timer.Interval=1000
 $timer.Add_Tick({
  if(-not $script:UpdateRequestActive){return}
  $p=$script:UpdateStatusPath
  if($p -and (Test-Path -LiteralPath $p)){
   try{
    $message=([IO.File]::ReadAllText($p)).Trim()
    if($message){
     switch -Wildcard ($message){
      'CHECKING' {$script:Status.Text='最新バージョンを確認しています…'}
      'DOWNLOADING' {$script:Status.Text='更新ファイルをダウンロードしています…'}
      'UPDATE_READY' {$script:Status.Text='更新の準備が完了しました。自動的に適用します。'}
      'UPDATE_DEFERRED' {$script:Status.Text='Excelが開いているため更新を保留しました。終了後に再試行します。'}
      'UP_TO_DATE' {$script:Status.Text='最新版を使用しています。'}
      'UPDATE_FAILED*' {$script:Status.Text='更新に失敗しました。ログを確認してください。'}
      default {$script:Status.Text='更新状況を確認しています…'}
     }
    }
    if(ShouldCloseSettingsForUpdate $script:UpdateRequestActive $message){
     $script:ShutdownReason='fresh-GitHub-update'
     WriteSettingsLog 'UPDATE_READY_CLOSE'
     $timer.Stop();$form.Close()
    }elseif($message -eq 'UP_TO_DATE' -or $message -eq 'UPDATE_DEFERRED' -or $message -like 'UPDATE_FAILED*'){
     $script:UpdateRequestActive=$false
     WriteSettingsLog 'UPDATE_FINISHED' $message
    }
   }catch{WriteSettingsLog 'STATUS_CHECK_ERROR' $_.Exception.Message}
  }
 })
 if($SelfTest){
  if($script:HotkeyFields.Count -ne 4 -or $form.Controls.Count -lt 15){throw 'Settings UI did not initialize completely'}
  Write-Output "SETTINGS_SELFTEST_OK version=$version controls=$($form.Controls.Count)"
  $form.Dispose()
  return
 }
 $timer.Start()
 $form.Add_Shown({
  WriteSettingsLog 'WINDOW_SHOWN'
  $form.Activate();$form.BringToFront()
 })
 $form.Add_FormClosing({
  WriteSettingsLog 'WINDOW_CLOSING' $script:ShutdownReason
 })
 WriteSettingsLog 'OPENING_DIALOG'
 [void]$form.ShowDialog()
 WriteSettingsLog 'DIALOG_RETURNED' $script:ShutdownReason
}catch{
 $details="起動エラー: "+$_.Exception.Message+" (line "+$_.InvocationInfo.ScriptLineNumber+")"
 WriteSettingsLog 'SCRIPT_EXCEPTION' $details
 try{Add-Content -LiteralPath $log -Value ("$(Get-Date -Format o) "+$details) -Encoding UTF8}catch{}
 try{[System.Windows.Forms.MessageBox]::Show($details+' 詳細はexcel-daily-settings.logを確認してください。','Excel Daily 設定エラー')|Out-Null}catch{}
 Write-Error $details
 exit 1
}
