$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding($false)
$dir=$env:ED_ROOT
$exe=Join-Path $dir 'excel-daily.exe'
function OwnAgents {
 @(Get-CimInstance Win32_Process -Filter "name='excel-daily.exe'" -ErrorAction Stop | Where-Object{$_.ExecutablePath -ieq $exe})
}
function Busy {
 if(Test-Path -LiteralPath (Join-Path $dir 'excel-daily-command.json')){return $true}
 if(Test-Path -LiteralPath (Join-Path $dir 'excel-daily-busy.lock')){
  return (@(OwnAgents).Count -gt 0)
 }
 return $false
}
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
  foreach($p in @(OwnAgents)){
   Stop-Process -Id $p.ProcessId -ErrorAction SilentlyContinue
   Wait-Process -Id $p.ProcessId -Timeout 10 -ErrorAction SilentlyContinue
  }
  Status
 }
 'restart' {
  if(Busy){throw 'Busy'}
  if(!(Test-Path -LiteralPath $exe)){throw 'Agent missing'}
  foreach($p in @(OwnAgents)){
   Stop-Process -Id $p.ProcessId -ErrorAction SilentlyContinue
   Wait-Process -Id $p.ProcessId -Timeout 10 -ErrorAction SilentlyContinue
  }
  Start-Sleep -Milliseconds 300
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
  if(Test-Path -LiteralPath $path){throw 'Busy'}
  $tmp=Join-Path $dir ('excel-daily-command-'+[guid]::NewGuid().ToString('N')+'.tmp')
  try{
   [IO.File]::WriteAllText($tmp,'{"action":"tomorrow"}',(New-Object Text.UTF8Encoding($false)))
   [IO.File]::Move($tmp,$path)
  }finally{
   if(Test-Path -LiteralPath $tmp){Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue}
  }
  Status
 }
 default {throw 'Invalid action'}
}
