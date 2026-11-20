param(
 [Parameter(Mandatory=$true)][string]$Executable,
 [Parameter(Mandatory=$true)][string]$Archive,
 [int]$PidToWait=0,
 [Parameter(Mandatory=$true)][string]$TargetVersion
)
$ErrorActionPreference='Stop'
$dir=Split-Path -Parent $Executable
$log=Join-Path $dir 'excel-daily-updater.log'
function Log([string]$msg){Add-Content -LiteralPath $log -Value ((Get-Date).ToString('s')+' '+$msg)}
$stage=Join-Path ([IO.Path]::GetTempPath()) ('excel-daily-extract-'+[guid]::NewGuid().ToString('N'))
try {
 if($TargetVersion -notmatch '^\d+\.\d+\.\d+$'){throw 'Invalid target version'}
 New-Item -ItemType Directory -Path $stage -Force|Out-Null
 Expand-Archive -LiteralPath $Archive -DestinationPath $stage -Force
 foreach($name in @('excel-daily.exe','excel-daily-settings.exe')){
  $src=Join-Path $stage $name
  if(!(Test-Path -LiteralPath $src) -or (Get-Item -LiteralPath $src).Length -lt 100KB){throw "Missing or invalid $name"}
 }
 if($PidToWait -gt 0){
  try{Wait-Process -Id $PidToWait -Timeout 60 -ErrorAction Stop}catch{
   if(Get-Process -Id $PidToWait -ErrorAction SilentlyContinue){throw 'Parent process still running'}
  }
 }
 # Never interrupt a running Excel write, including updates initiated in Settings.
 $busy=Join-Path $dir 'excel-daily-busy.lock'
 for($i=0;$i -lt 120;$i++){
  if(!(Test-Path -LiteralPath $busy)){break}
  Start-Sleep -Seconds 1
 }
 if(Test-Path -LiteralPath $busy){throw 'Excel update active; retry later'}
 # Only stop this installation's processes, not another Excel Daily copy.
 foreach($name in @('excel-daily-settings.exe','excel-daily.exe')){
  $target=Join-Path $dir $name
  $live=@(Get-CimInstance Win32_Process -Filter ("name='"+$name+"'") -ErrorAction SilentlyContinue |
    Where-Object{$_.ExecutablePath -ieq $target})
  foreach($p in $live){Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop}
 }
 Start-Sleep -Milliseconds 600
 foreach($name in @('excel-daily.exe','excel-daily-settings.exe')){
  $dst=Join-Path $dir $name
  if(Test-Path -LiteralPath $dst){Copy-Item -LiteralPath $dst -Destination ($dst+'.bak') -Force}
  Copy-Item -LiteralPath (Join-Path $stage $name) -Destination $dst -Force
 }
 # Excel may hold the XLL open. Never fail an Agent update or interrupt Excel
 # because its Ribbon Add-in cannot be replaced until Excel is closed.
 foreach($name in @('ExcelDailyRibbon.xll','ExcelDailyRibbon64.xll','install-ribbon.ps1')){
  $src=Join-Path $stage $name
  if(!(Test-Path -LiteralPath $src)){continue}
  $dst=Join-Path $dir $name
  try{
   Copy-Item -LiteralPath $src -Destination $dst -Force -ErrorAction Stop
  }catch{
   Log ('RIBBON_COPY_DEFERRED '+$name+'; close Excel and reinstall from the ZIP if required: '+$_.Exception.Message)
  }
 }
 # settings.json, Excel workbooks, excel-daily-state.json and device id are preserved.
 $exe=Join-Path $dir 'excel-daily.exe'
 $agent=Start-Process -FilePath $exe -WorkingDirectory $dir -PassThru
 Start-Sleep -Seconds 2
 if($agent.HasExited){throw "Updated agent exited with code $($agent.ExitCode)"}
 Log ('UPDATE_SUCCESS version='+$TargetVersion+' pid='+$agent.Id)
}catch{
 Log ('UPDATE_FAILED '+$_.Exception.Message+'; backups left for manual recovery')
 exit 1
}finally{
 Remove-Item -LiteralPath $stage -Force -Recurse -ErrorAction SilentlyContinue
 Remove-Item -LiteralPath (Split-Path -Parent $Archive) -Force -Recurse -ErrorAction SilentlyContinue
}
