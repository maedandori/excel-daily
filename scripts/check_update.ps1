param([string]$Executable,[string]$CurrentVersion,[string]$Repo,[int]$PidToWait)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
try {
 if ($Repo -notmatch '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') { throw 'Invalid repository' }
 [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
 $headers=@{'User-Agent'='excel-daily-updater';'Accept'='application/vnd.github+json'}
 $release=Invoke-RestMethod -Uri ("https://api.github.com/repos/"+$Repo+"/releases/latest") -Headers $headers -TimeoutSec 20
 $version=([string]$release.tag_name).TrimStart('v')
 if ($version -notmatch '^\d+\.\d+\.\d+$' -or [version]$version -le [version]$CurrentVersion) { exit 0 }
 $exeAsset=@($release.assets | Where-Object {$_.name -eq 'excel-daily.exe'}) | Select-Object -First 1
 $shaAsset=@($release.assets | Where-Object {$_.name -eq 'excel-daily.exe.sha256'}) | Select-Object -First 1
 if (!$exeAsset -or !$shaAsset) { throw 'Missing release assets' }
 $stage=Join-Path ([IO.Path]::GetTempPath()) ('excel-daily-update-'+[guid]::NewGuid().ToString('N'))
 New-Item -Path $stage -ItemType Directory -Force | Out-Null
 $target=Join-Path $stage 'excel-daily.exe'
 $checksum=Join-Path $stage 'excel-daily.exe.sha256'
 Invoke-WebRequest -UseBasicParsing -Uri $exeAsset.browser_download_url -OutFile $target -TimeoutSec 90
 Invoke-WebRequest -UseBasicParsing -Uri $shaAsset.browser_download_url -OutFile $checksum -TimeoutSec 30
 $expected=(([IO.File]::ReadAllText($checksum)).Trim() -split '\s+')[0]
 if ($expected -notmatch '^[A-Fa-f0-9]{64}$') { throw 'Malformed SHA256' }
 if ((Get-FileHash -Algorithm SHA256 -LiteralPath $target).Hash -ine $expected) { throw 'SHA256 mismatch' }
 $installer=Join-Path (Split-Path -Parent $Executable) 'apply_update.ps1'
 if (!(Test-Path -LiteralPath $installer)) { throw 'Updater helper missing' }
 $arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $installer + '"'),'-Executable',('"' + $Executable + '"'),'-StagedExe',('"' + $target + '"'),'-PidToWait',[string]$PidToWait)
 Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -WindowStyle Hidden | Out-Null
 Write-Output 'UPDATE_READY'
} catch { Write-Error ('Auto-update: '+$_); exit 1 }
