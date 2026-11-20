# 安全な更新のため、Excel と Excel-daily の稼働状況を確認する。
# 戻り値は表示文ではなく、機械判定用の固定コード。
function Get-ExcelDailyUpdateBlocker {
 param([Parameter(Mandatory=$true)][string]$Root,[Nullable[bool]]$ExcelOpen=$null)
 if($null -eq $ExcelOpen){
  $ExcelOpen=(@(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue).Count -gt 0)
 }
 if($ExcelOpen){return 'EXCEL_OPEN'}
 if((Test-Path -LiteralPath (Join-Path $Root 'excel-daily-busy.lock')) -or
    (Test-Path -LiteralPath (Join-Path $Root 'excel-daily-command.json'))){return 'EXCEL_DAILY_BUSY'}
 return ''
}
