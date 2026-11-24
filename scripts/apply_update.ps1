param(
 [Parameter(Mandatory=$true)][string]$Executable,
 [Parameter(Mandatory=$true)][string]$Archive,
 [int]$PidToWait=0,
 [Parameter(Mandatory=$true)][string]$TargetVersion
)
$ErrorActionPreference='Stop'
$dir=Split-Path -Parent $Executable
$log=Join-Path $dir 'excel-daily-updater.log'
. (Join-Path $PSScriptRoot 'update_safety.ps1')
function Log([string]$msg){Add-Content -LiteralPath $log -Value ((Get-Date).ToString('s')+' '+$msg)}
function UpdateStatus([string]$state,[string]$message){
 try{
  $json=@{state=$state;message=$message;updated_ms=[DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()}|ConvertTo-Json -Compress
  [IO.File]::WriteAllText((Join-Path $dir 'excel-daily-status.json'),$json,([Text.UTF8Encoding]::new($false)))
 }catch{Log ('STATUS_WRITE_FAILED '+$_.Exception.Message)}
}
function RestartAgent {
 $agent=Start-Process -FilePath $Executable -WorkingDirectory $dir -PassThru -ErrorAction Stop
 Start-Sleep -Seconds 2
 if($agent.HasExited){throw "Excel Dailyの再起動に失敗しました (終了コード: $($agent.ExitCode))"}
 return $agent.Id
}
$stage=Join-Path ([IO.Path]::GetTempPath()) ('excel-daily-extract-'+[guid]::NewGuid().ToString('N'))
$backup=Join-Path ([IO.Path]::GetTempPath()) ('excel-daily-backup-'+[guid]::NewGuid().ToString('N'))
$stopped=$false
$copyStarted=$false
$copied=@()
try {
 if($TargetVersion -notmatch '^\d+\.\d+\.\d+$'){throw '更新バージョンが正しくありません。'}
 New-Item -ItemType Directory -Path $stage -Force|Out-Null
 Expand-Archive -LiteralPath $Archive -DestinationPath $stage -Force
 # Validate every binary BEFORE stopping the running Agent.
 $required=@('excel-daily.exe','excel-daily-settings.exe','ExcelDailyRibbon.xll','ExcelDailyRibbon64.xll','install-ribbon.ps1')
 foreach($name in $required){
  $src=Join-Path $stage $name
  $minimum=if($name -like '*.exe' -or $name -like '*.xll'){100KB}else{100}
  if(!(Test-Path -LiteralPath $src) -or (Get-Item -LiteralPath $src).Length -lt $minimum){
   throw "更新ファイルが不足しています: $name"
  }
 }
 if($PidToWait -gt 0){
  try{Wait-Process -Id $PidToWait -Timeout 60 -ErrorAction Stop}catch{
   if(Get-Process -Id $PidToWait -ErrorAction SilentlyContinue){throw '更新元の終了を確認できません。'}
  }
 }
 # Wait for an active workbook write, not for Excel itself to close.
 $reason=Get-ExcelDailyUpdateBlocker -Root $dir
 if($reason){
  Log ('UPDATE_DEFERRED reason='+$reason)
  UpdateStatus 'deferred' '更新処理中のため保留しました。完了後に再試行します。'
  $id=RestartAgent
  Log ('OLD_AGENT_RESTARTED pid='+$id)
  exit 0
 }
 # Stop only this installation's binaries. Never close Excel or another Agent.
 foreach($name in @('excel-daily-settings.exe','excel-daily.exe')){
  $target=Join-Path $dir $name
  $live=@(Get-CimInstance Win32_Process -Filter ("name='"+$name+"'") -ErrorAction SilentlyContinue |
    Where-Object{$_.ExecutablePath -ieq $target})
  foreach($p in $live){Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop}
 }
 $stopped=$true
 Start-Sleep -Milliseconds 500
 # Check again to avoid installing when Excel reopened in the meantime.
 $reason=Get-ExcelDailyUpdateBlocker -Root $dir
 if($reason){
  Log ('UPDATE_DEFERRED_AFTER_STOP reason='+$reason)
  UpdateStatus 'deferred' '更新処理中のため保留しました。完了後に再試行します。'
  $id=RestartAgent
  $stopped=$false
  Log ('OLD_AGENT_RESTARTED pid='+$id)
  exit 0
 }
 New-Item -ItemType Directory -Path $backup -Force | Out-Null
 $copyStarted=$true
 # Keep the running Excel session intact. Stage XLLs separately if loaded.
 $ribbonPending=$false
 foreach($name in $required){
  # Always stage XLLs; the agent can run with the previous compatible Ribbon.
  # This also avoids a race when Excel starts during installation.
  $destName=if($name -like '*.xll'){$ribbonPending=$true;$name+'.pending'}else{$name}
  $dst=Join-Path $dir $destName
  $hadOld=Test-Path -LiteralPath $dst
  if($hadOld){Copy-Item -LiteralPath $dst -Destination (Join-Path $backup $destName) -Force -ErrorAction Stop}
  $copied+=@{path=$dst;name=$destName;hadOld=$hadOld}
  Copy-Item -LiteralPath (Join-Path $stage $name) -Destination $dst -Force -ErrorAction Stop
 }
 $id=RestartAgent
 $stopped=$false
 Log ('UPDATE_SUCCESS version='+$TargetVersion+' pid='+$id)
 UpdateStatus 'success' ('更新が完了しました。バージョン '+$TargetVersion+'。Excel Dailyを再起動しました。リボンはExcel終了後に更新します。')
}catch{
 Log ('UPDATE_FAILED '+$_.Exception.Message)
 if($copyStarted){
  foreach($item in @($copied)){
   try{
    $prior=Join-Path $backup $item.name
    if($item.hadOld -and (Test-Path -LiteralPath $prior)){
     Copy-Item -LiteralPath $prior -Destination $item.path -Force -ErrorAction Stop
    } elseif(-not $item.hadOld -and (Test-Path -LiteralPath $item.path)){
     Remove-Item -LiteralPath $item.path -Force -ErrorAction Stop
    }
   }catch{Log ('ROLLBACK_FAILED '+$item.name+': '+$_.Exception.Message)}
  }
 }
 if($stopped){
  try{$id=RestartAgent;Log ('ROLLBACK_AGENT_RESTARTED pid='+$id)}
  catch{Log ('AGENT_RESTART_FAILED '+$_.Exception.Message)}
 }
 UpdateStatus 'error' '自動更新に失敗しました。詳細はexcel-daily-updater.logを確認してください。'
 exit 1
}finally{
 Remove-Item -LiteralPath $stage -Force -Recurse -ErrorAction SilentlyContinue
 Remove-Item -LiteralPath $backup -Force -Recurse -ErrorAction SilentlyContinue
 if($Archive -and (Test-Path -LiteralPath $Archive)){
  $parent=Split-Path -Parent $Archive
  if((Split-Path -Leaf $parent) -like 'excel-daily-upgrade-*'){
   Remove-Item -LiteralPath $parent -Force -Recurse -ErrorAction SilentlyContinue
  }
 }
}
