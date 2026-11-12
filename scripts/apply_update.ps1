param([string]$Executable,[string]$StagedExe,[int]$PidToWait)
$ErrorActionPreference='Stop'
$log=Join-Path (Split-Path -Parent $Executable) 'excel-daily-updater.log'
function Log($s) { Add-Content -LiteralPath $log -Value ((Get-Date).ToString('s')+' '+$s) }
try {
 try { Wait-Process -Id $PidToWait -Timeout 45 -ErrorAction Stop } catch {
   if (Get-Process -Id $PidToWait -ErrorAction SilentlyContinue) { throw 'Old process did not exit' }
 }
 $backup=$Executable+'.bak'
 Copy-Item -LiteralPath $Executable -Destination $backup -Force
 try {
   Copy-Item -LiteralPath $StagedExe -Destination $Executable -Force
   $p=Start-Process -FilePath $Executable -WorkingDirectory (Split-Path -Parent $Executable) -PassThru
   Log ('Updated and restarted PID='+$p.Id)
 } catch {
   Copy-Item -LiteralPath $backup -Destination $Executable -Force
   Log ('Restored backup: '+$_)
   throw
 }
} catch { Log ('FAILED: '+$_); exit 1 } finally {
 Remove-Item -LiteralPath (Split-Path -Parent $StagedExe) -Recurse -Force -ErrorAction SilentlyContinue
}
