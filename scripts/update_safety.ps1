# 安全な更新のため、Excel と Excel-daily の稼働状況を確認する。
# 戻り値は表示文ではなく、機械判定用の固定コード。
function Get-ExcelDailyUpdateBlocker {
 param([Parameter(Mandatory=$true)][string]$Root,[Nullable[bool]]$ExcelOpen=$null)
 if($null -eq $ExcelOpen){
  $ExcelOpen=(@(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue).Count -gt 0)
 }
 if((Test-Path -LiteralPath (Join-Path $Root 'excel-daily-busy.lock')) -or
    (Test-Path -LiteralPath (Join-Path $Root 'excel-daily-command.json'))){return 'EXCEL_DAILY_BUSY'}
 return ''
}

# Replace pending Ribbon binaries only when no Excel process is using them.
function Complete-ExcelDailyPendingRibbon {
 param([Parameter(Mandatory=$true)][string]$Root)
 if(@(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue).Count -gt 0){return $false}
 foreach($name in @('ExcelDailyRibbon.xll','ExcelDailyRibbon64.xll')){
  $pending=Join-Path $Root ($name+'.pending')
  if(!(Test-Path -LiteralPath $pending)){continue}
  try{
   Copy-Item -LiteralPath $pending -Destination (Join-Path $Root $name) -Force -ErrorAction Stop
   Remove-Item -LiteralPath $pending -Force -ErrorAction Stop
  }catch{return $false}
 }
 return $true
}
