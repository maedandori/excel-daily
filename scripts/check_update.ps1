param(
 [Parameter(Mandatory=$true)][string]$Executable,
 [Parameter(Mandatory=$true)][string]$CurrentVersion,
 [Parameter(Mandatory=$true)][string]$Repo,
 [int]$PidToWait=0,
 [string]$StatusFile=''
)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
. (Join-Path $PSScriptRoot 'update_safety.ps1')
$stage='';$handedOff=$false
function Status([string]$s){if($StatusFile){[IO.File]::WriteAllText($StatusFile,$s,(New-Object Text.UTF8Encoding($false)))}}
try {
 $dir=Split-Path -Parent $Executable
 $blocker=Get-ExcelDailyUpdateBlocker -Root $dir
 if($blocker){Status 'UPDATE_DEFERRED';Write-Output 'UPDATE_DEFERRED';exit 0}
 $null=Complete-ExcelDailyPendingRibbon -Root $dir
 if($Repo -ne 'maedandori/excel-daily'){throw '更新元のリポジトリが正しくありません。'}
 if($CurrentVersion -notmatch '^\d+\.\d+\.\d+$'){throw '現在のバージョンが正しくありません。'}
 [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
 $headers=@{'User-Agent'='excel-daily-updater';'Accept'='application/vnd.github+json'}
 Status 'CHECKING'
 $release=Invoke-RestMethod -Uri ('https://api.github.com/repos/'+$Repo+'/releases/latest') -Headers $headers -TimeoutSec 20
 $version=([string]$release.tag_name) -replace '^(ed|v)',''
 if($version -notmatch '^\d+\.\d+\.\d+$'){throw '配信されたバージョンが正しくありません。'}
 if([version]$version -le [version]$CurrentVersion){Status 'UP_TO_DATE';exit 0}
 $zipAsset=@($release.assets|Where-Object{$_.name -eq 'excel-daily-windows.zip'})|Select-Object -First 1
 $shaAsset=@($release.assets|Where-Object{$_.name -eq 'excel-daily-windows.zip.sha256'})|Select-Object -First 1
 if(!$zipAsset -or !$shaAsset){throw '更新ファイルまたはチェックサムが見つかりません。'}
 $stage=Join-Path ([IO.Path]::GetTempPath()) ('excel-daily-upgrade-'+[guid]::NewGuid().ToString('N'))
 New-Item -ItemType Directory -Path $stage -Force|Out-Null
 $zip=Join-Path $stage 'excel-daily-windows.zip'
 $checksum=Join-Path $stage 'excel-daily-windows.zip.sha256'
 Status 'DOWNLOADING'
 Invoke-WebRequest -UseBasicParsing -Uri $zipAsset.browser_download_url -OutFile $zip -TimeoutSec 120
 Invoke-WebRequest -UseBasicParsing -Uri $shaAsset.browser_download_url -OutFile $checksum -TimeoutSec 30
 $expected=([IO.File]::ReadAllText($checksum).Trim() -split '\s+')[0]
 if($expected -notmatch '^[A-Fa-f0-9]{64}$'){throw '更新ファイルのチェックサムが不正です。'}
 if((Get-FileHash -Algorithm SHA256 -LiteralPath $zip).Hash -ine $expected){throw '更新ファイルの検証に失敗しました。'}
 if((Get-Item -LiteralPath $zip).Length -lt 100KB -or (Get-Item -LiteralPath $zip).Length -gt 40MB){throw '更新ファイルのサイズが不正です。'}
 $blocker=Get-ExcelDailyUpdateBlocker -Root $dir
 if($blocker){Status 'UPDATE_DEFERRED';Write-Output 'UPDATE_DEFERRED';exit 0}
 $installer=Join-Path (Split-Path -Parent $Executable) 'apply_update.ps1'
 if(!(Test-Path -LiteralPath $installer)){throw '更新用プログラムが見つかりません。'}
 $args=@('-NoProfile','-NonInteractive','-WindowStyle','Hidden','-ExecutionPolicy','Bypass','-File',('"' + $installer + '"'),
  '-Executable',('"' + $Executable + '"'),'-Archive',('"' + $zip + '"'),
  '-PidToWait',[string]$PidToWait,'-TargetVersion',$version)
 Start-Process -FilePath powershell.exe -ArgumentList $args -WindowStyle Hidden|Out-Null
 $handedOff=$true
 Status 'UPDATE_READY'
 Write-Output 'UPDATE_READY'
}catch{
 Status ('UPDATE_FAILED '+$_.Exception.Message)
 Write-Error ('Excel Daily アップデートエラー: '+$_.Exception.Message)
 exit 1
}finally{
 if($stage -and -not $handedOff -and (Test-Path -LiteralPath $stage)){
  Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
 }
}
