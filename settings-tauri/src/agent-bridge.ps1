$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding($false)
$dir=$env:ED_ROOT
$exe=Join-Path $dir 'excel-daily.exe'
function OwnAgents {
 @(Get-CimInstance Win32_Process -Filter "name='excel-daily.exe'" -ErrorAction Stop | Where-Object{$_.ExecutablePath -ieq $exe})
}
function Busy { (Test-Path -LiteralPath (Join-Path $dir 'excel-daily-busy.lock')) -or (Test-Path -LiteralPath (Join-Path $dir 'excel-daily-command.json')) }
function Status {
 $live=@(OwnAgents)
 $version=''
 try{
  $v=Get-Content -LiteralPath (Join-Path $dir 'excel-daily-version.json') -Raw -Encoding UTF8|ConvertFrom-Json
  if($live.ProcessId -contains [int]$v.pid){$version=[string]$v.version}
 }catch{}
 @{running=($live.Count -gt 0);version=$version;busy=[bool](Busy)}|ConvertTo-Json -Compress
}
switch($env:ED_ACTION){
 'status' { Status }
 'stop' {
  if(Busy){throw 'Busy'}
  foreach($p in @(OwnAgents)){Stop-Process -Id $p.ProcessId -ErrorAction Stop}
  Status
 }
 'restart' {
  if(Busy){throw 'Busy'}
  if(!(Test-Path -LiteralPath $exe)){throw 'Agent missing'}
  foreach($p in @(OwnAgents)){Stop-Process -Id $p.ProcessId -ErrorAction Stop}
  Start-Sleep -Milliseconds 700
  $p=Start-Process -FilePath $exe -WorkingDirectory $dir -PassThru
  Start-Sleep -Milliseconds 900
  if($p.HasExited){throw 'Agent exited'}
  Status
 }
 'tomorrow' {
  if(Busy){throw 'Busy'}
  $live=@(OwnAgents)
  $v=Get-Content -LiteralPath (Join-Path $dir 'excel-daily-version.json') -Raw -Encoding UTF8|ConvertFrom-Json
  if(!($live.ProcessId -contains [int]$v.pid) -or [version]$v.version -lt [version]'0.3.2'){throw 'Agent not running'}
  $path=Join-Path $dir 'excel-daily-command.json'
  $stream=[IO.File]::Open($path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
  try{$bytes=[Text.Encoding]::UTF8.GetBytes('{"action":"tomorrow"}');$stream.Write($bytes,0,$bytes.Length)}finally{$stream.Dispose()}
  Status
 }
 default {throw 'Invalid action'}
}
