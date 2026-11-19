param(
 [Parameter(Mandatory=$true)][string]$Executable,
 [Parameter(Mandatory=$true)][string]$CurrentVersion,
 [Parameter(Mandatory=$true)][string]$Repo,
 [int]$PidToWait=0,
 [string]$StatusFile=''
)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
function Status([string]$s){if($StatusFile){[IO.File]::WriteAllText($StatusFile,$s,(New-Object Text.UTF8Encoding($false)))}}
try {
 if($Repo -ne 'maedandori/excel-daily'){throw 'Unexpected updater repository'}
 if($CurrentVersion -notmatch '^\d+\.\d+\.\d+$'){throw 'Invalid installed version'}
 [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
 $headers=@{'User-Agent'='excel-daily-updater';'Accept'='application/vnd.github+json'}
 Status 'Checking latest release...'
 $release=Invoke-RestMethod -Uri ('https://api.github.com/repos/'+$Repo+'/releases/latest') -Headers $headers -TimeoutSec 20
 $version=([string]$release.tag_name).TrimStart('v')
 if($version -notmatch '^\d+\.\d+\.\d+$'){throw 'Invalid release version'}
 if([version]$version -le [version]$CurrentVersion){Status 'UP_TO_DATE';exit 0}
 $zipAsset=@($release.assets|Where-Object{$_.name -eq 'excel-daily-windows.zip'})|Select-Object -First 1
 $shaAsset=@($release.assets|Where-Object{$_.name -eq 'excel-daily-windows.zip.sha256'})|Select-Object -First 1
 if(!$zipAsset -or !$shaAsset){throw 'Release missing portable ZIP/checksum'}
 $stage=Join-Path ([IO.Path]::GetTempPath()) ('excel-daily-upgrade-'+[guid]::NewGuid().ToString('N'))
 New-Item -ItemType Directory -Path $stage -Force|Out-Null
 $zip=Join-Path $stage 'excel-daily-windows.zip'
 $checksum=Join-Path $stage 'excel-daily-windows.zip.sha256'
 Status 'DOWNLOADING'
 Invoke-WebRequest -UseBasicParsing -Uri $zipAsset.browser_download_url -OutFile $zip -TimeoutSec 120
 Invoke-WebRequest -UseBasicParsing -Uri $shaAsset.browser_download_url -OutFile $checksum -TimeoutSec 30
 $expected=([IO.File]::ReadAllText($checksum).Trim() -split '\s+')[0]
 if($expected -notmatch '^[A-Fa-f0-9]{64}$'){throw 'Invalid checksum'}
 if((Get-FileHash -Algorithm SHA256 -LiteralPath $zip).Hash -ine $expected){throw 'Release ZIP checksum mismatch'}
 if((Get-Item -LiteralPath $zip).Length -lt 100KB -or (Get-Item -LiteralPath $zip).Length -gt 40MB){throw 'ZIP size invalid'}
 $installer=Join-Path (Split-Path -Parent $Executable) 'apply_update.ps1'
 if(!(Test-Path -LiteralPath $installer)){throw 'Update helper missing'}
 $args=@('-NoProfile','-NonInteractive','-WindowStyle','Hidden','-ExecutionPolicy','Bypass','-File',('"' + $installer + '"'),
  '-Executable',('"' + $Executable + '"'),'-Archive',('"' + $zip + '"'),
  '-PidToWait',[string]$PidToWait,'-TargetVersion',$version)
 Start-Process -FilePath powershell.exe -ArgumentList $args -WindowStyle Hidden|Out-Null
 Status 'UPDATE_READY'
 Write-Output 'UPDATE_READY'
}catch{
 Status ('UPDATE_FAILED '+$_.Exception.Message)
 Write-Error ('Excel Daily updater: '+$_.Exception.Message)
 exit 1
}
