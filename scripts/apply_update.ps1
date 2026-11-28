param(
 [Parameter(Mandatory=$true)][string]$Executable,
 [Parameter(Mandatory=$true)][string]$Archive,
 [int]$PidToWait=0,
 [Parameter(Mandatory=$true)][string]$TargetVersion,
 [string]$ExpectedSha256=''
)
$ErrorActionPreference='Stop'
$dir=Split-Path -Parent $Executable
$log=Join-Path $dir 'excel-daily-updater.log'
. (Join-Path $PSScriptRoot 'update_safety.ps1')
function Log([string]$msg){Add-Content -LiteralPath $log -Value ((Get-Date).ToString('s')+' '+$msg) -Encoding UTF8}
function UpdateStatus([string]$state,[string]$message){
 try{
  $json=@{state=$state;message=$message;updated_ms=[DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()}|ConvertTo-Json -Compress
  $statusPath=Join-Path $dir 'excel-daily-status.json'
  $tmp=Join-Path $dir ('excel-daily-status-'+[guid]::NewGuid().ToString('N')+'.tmp')
  [IO.File]::WriteAllText($tmp,$json,([Text.UTF8Encoding]::new($false)))
  [IO.File]::Copy($tmp,$statusPath,$true)
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
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
 if(!(Test-Path -LiteralPath $Archive)){throw '更新ZIPファイルが見つかりません。'}
 $archiveLen=(Get-Item -LiteralPath $Archive).Length
 if($archiveLen -lt 100KB -or $archiveLen -gt 40MB){throw '更新ZIPファイルのサイズが不正です。'}
 if($ExpectedSha256){
  if($ExpectedSha256 -notmatch '^[A-Fa-f0-9]{64}$'){throw 'チェックサムの形式が不正です。'}
  if((Get-FileHash -Algorithm SHA256 -LiteralPath $Archive).Hash -ine $ExpectedSha256){throw '更新ファイルのSHA256検証に失敗しました。'}
 }
 if(Test-Path -LiteralPath $stage){throw 'Staging directory already exists'}
 New-Item -ItemType Directory -Path $stage | Out-Null
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $stageRoot=[IO.Path]::GetFullPath($stage).TrimEnd('\')+'\'
 $zip=[IO.Compression.ZipFile]::OpenRead($Archive)
 try{
  if($zip.Entries.Count -gt 200){throw 'Too many ZIP entries'}
  $totalBytes=0L
  foreach($entry in $zip.Entries){
   $name=[string]$entry.FullName
   if($name -match '(^|[\\/])\.\.([\\/]|$)|[\\/]|^[A-Za-z]:'){throw "Unsafe ZIP entry: $name"}
   if($name -notmatch '^(?:[A-Za-z0-9_.-]+\.ps1|excel-daily\.exe|excel-daily-settings\.exe|ExcelDailyRibbon(?:64)?\.xll|ribbon-version\.json|settings\.json|events\.example\.json|README\.md)$'){
    throw "Unexpected ZIP entry: $name"
   }
   $totalBytes+=$entry.Length
   if($totalBytes -gt 150MB){throw 'ZIP uncompressed size too large'}
   $destFull=[IO.Path]::GetFullPath((Join-Path $stage $name))
   if(-not $destFull.StartsWith($stageRoot,[StringComparison]::OrdinalIgnoreCase)){throw "Unsafe ZIP destination: $name"}
   $inStream=$entry.Open()
   $outStream=[IO.File]::Create($destFull)
   try{$inStream.CopyTo($outStream)}finally{$inStream.Dispose();$outStream.Dispose()}
  }
 }finally{$zip.Dispose()}
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
  foreach($p in $live){
   Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
   Wait-Process -Id $p.ProcessId -Timeout 15 -ErrorAction SilentlyContinue
  }
 }
 $stopped=$true
 Start-Sleep -Milliseconds 300
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
 $ribbonManifest=Join-Path $stage 'ribbon-version.json'
 $ribbonPlan=Get-ExcelDailyRibbonPlan -Root $dir -Stage $stage
 Log ('RIBBON_PLAN '+$ribbonPlan.reason)
 foreach($name in $required){
  if($name -like '*.xll' -and -not $ribbonPlan.replace){continue}
  # Always stage XLLs; the agent can run with the previous compatible Ribbon.
  # This also avoids a race when Excel starts during installation.
  $destName=if($name -like '*.xll'){$ribbonPending=$true;$name+'.pending'}else{$name}
  $dst=Join-Path $dir $destName
  $hadOld=Test-Path -LiteralPath $dst
  if($hadOld){Copy-Item -LiteralPath $dst -Destination (Join-Path $backup $destName) -Force -ErrorAction Stop}
  $copied+=@{path=$dst;name=$destName;hadOld=$hadOld}
  Copy-Item -LiteralPath (Join-Path $stage $name) -Destination $dst -Force -ErrorAction Stop
 }
 if($ribbonPlan.replace -and (Test-Path -LiteralPath $ribbonManifest)){
  $name='ribbon-version.json.pending'
  $dst=Join-Path $dir $name
  $hadOld=Test-Path -LiteralPath $dst
  if($hadOld){Copy-Item -LiteralPath $dst -Destination (Join-Path $backup $name) -Force -ErrorAction Stop}
  $copied+=@{path=$dst;name=$name;hadOld=$hadOld}
  Copy-Item -LiteralPath $ribbonManifest -Destination $dst -Force -ErrorAction Stop
 } elseif($ribbonPlan.reason -eq 'same_hash' -and (Test-Path -LiteralPath $ribbonManifest)){
  $name='ribbon-version.json'
  $dst=Join-Path $dir $name
  $hadOld=Test-Path -LiteralPath $dst
  if($hadOld){Copy-Item -LiteralPath $dst -Destination (Join-Path $backup $name) -Force -ErrorAction Stop}
  $copied+=@{path=$dst;name=$name;hadOld=$hadOld}
  Copy-Item -LiteralPath $ribbonManifest -Destination $dst -Force -ErrorAction Stop
 }
 $id=RestartAgent
 $stopped=$false
 Log ('UPDATE_SUCCESS version='+$TargetVersion+' pid='+$id)
 $message='更新が完了しました。バージョン '+$TargetVersion+'。Excel Dailyを再起動しました。'
 if($ribbonPending){$message+='リボンはExcel終了後に更新します。'}else{$message+='リボンの更新は不要です。'}
 UpdateStatus 'success' $message
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
  $tempRoot=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
  $parent=[IO.Path]::GetFullPath((Split-Path -Parent $Archive))
  if($parent.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Leaf $parent) -like 'excel-daily-upgrade-*'){
   Remove-Item -LiteralPath $parent -Force -Recurse -ErrorAction SilentlyContinue
  }
 }
}
