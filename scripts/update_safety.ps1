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
 # Only publish the revision after both pending binaries are installed and verified.
 $manifest=Join-Path $Root 'ribbon-version.json.pending'
 if(Test-Path -LiteralPath $manifest){
  try{
   $meta=Get-Content -LiteralPath $manifest -Raw -Encoding UTF8 | ConvertFrom-Json
   foreach($name in @('ExcelDailyRibbon.xll','ExcelDailyRibbon64.xll')){
    if((Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $Root $name)).Hash -ine [string]$meta.files.PSObject.Properties[$name].Value){return $false}
   }
   Copy-Item -LiteralPath $manifest -Destination (Join-Path $Root 'ribbon-version.json') -Force -ErrorAction Stop
   Remove-Item -LiteralPath $manifest -Force -ErrorAction Stop
  }catch{return $false}
 }
 return $true
}

function Get-ExcelDailyRibbonPlan {
 param([Parameter(Mandatory=$true)][string]$Root,[Parameter(Mandatory=$true)][string]$Stage)
 $names=@('ExcelDailyRibbon.xll','ExcelDailyRibbon64.xll')
 $incoming=$null
 $manifest=Join-Path $Stage 'ribbon-version.json'
 if(Test-Path -LiteralPath $manifest){
  $incoming=Get-Content -LiteralPath $manifest -Raw -Encoding UTF8 | ConvertFrom-Json
  if([string]$incoming.revision -notmatch '^[a-f0-9]{40}$'){throw 'Invalid Ribbon revision'}
  foreach($name in $names){
   $expected=[string]$incoming.files.PSObject.Properties[$name].Value
   if($expected -notmatch '^[a-f0-9]{64}$' -or (Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $Stage $name)).Hash -ine $expected){throw 'Ribbon checksum mismatch'}
  }
 }
 # Prefer a complete already-staged pair; never trust revision alone if files are missing or corrupt.
 $pendingManifest=Join-Path $Root 'ribbon-version.json.pending'
 $installedManifest=Join-Path $Root 'ribbon-version.json'
 foreach($candidate in @(@{path=$pendingManifest;suffix='.pending'},@{path=$installedManifest;suffix=''})){
  if(!$incoming -or !(Test-Path -LiteralPath $candidate.path)){continue}
  try{
   $current=Get-Content -LiteralPath $candidate.path -Raw -Encoding UTF8 | ConvertFrom-Json
   if($current.revision -ne $incoming.revision){continue}
   # An installed revision cannot override an unrelated staged upgrade.
   if(!$candidate.suffix -and @($names|Where-Object{Test-Path -LiteralPath (Join-Path $Root ($_+'.pending'))}).Count){continue}
   $valid=$true
   foreach($name in $names){
    $actualPath=Join-Path $Root ($name+$candidate.suffix)
    if(!(Test-Path -LiteralPath $actualPath) -or (Get-FileHash -Algorithm SHA256 -LiteralPath $actualPath).Hash -ine [string]$current.files.PSObject.Properties[$name].Value){$valid=$false;break}
   }
   if($valid){return @{replace=$false;reason='same_revision'}}
  }catch{}
 }
 $same=$true
 foreach($name in $names){
  $actualPath=Join-Path $Root $name
  if((Test-Path -LiteralPath ($actualPath+'.pending')) -or !(Test-Path -LiteralPath $actualPath) -or
      (Get-FileHash -Algorithm SHA256 -LiteralPath $actualPath).Hash -ine (Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $Stage $name)).Hash){$same=$false;break}
 }
 if($same){return @{replace=$false;reason='same_hash'}}
 return @{replace=$true;reason='changed_or_untracked'}
}
