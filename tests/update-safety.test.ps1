$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '..\scripts\update_safety.ps1')
$dir=Join-Path ([IO.Path]::GetTempPath()) ('excel-daily-update-guard-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir -Force|Out-Null
try{
 if((Get-ExcelDailyUpdateBlocker -Root $dir -ExcelOpen $true) -ne 'EXCEL_OPEN'){throw 'Excel running should defer update'}
 if((Get-ExcelDailyUpdateBlocker -Root $dir -ExcelOpen $false)){throw 'Idle should permit update'}
 [IO.File]::WriteAllText((Join-Path $dir 'excel-daily-command.json'),'{}')
 if((Get-ExcelDailyUpdateBlocker -Root $dir -ExcelOpen $false) -ne 'EXCEL_DAILY_BUSY'){throw 'Pending request should defer'}
 Remove-Item (Join-Path $dir 'excel-daily-command.json')
 [IO.File]::WriteAllText((Join-Path $dir 'excel-daily-busy.lock'),'123')
 if((Get-ExcelDailyUpdateBlocker -Root $dir -ExcelOpen $false) -ne 'EXCEL_DAILY_BUSY'){throw 'Workbook write should defer'}
 if((Get-ExcelDailyUpdateBlocker -Root $dir -ExcelOpen $true) -ne 'EXCEL_OPEN'){throw 'Excel-open priority wrong'}
 $source=Get-Content (Join-Path $PSScriptRoot '..\scripts\check_update.ps1') -Raw -Encoding UTF8
 $apply=Get-Content (Join-Path $PSScriptRoot '..\scripts\apply_update.ps1') -Raw -Encoding UTF8
 if($source.IndexOf('Get-ExcelDailyUpdateBlocker') -gt $source.IndexOf('Invoke-RestMethod')){throw 'Update check must short-circuit before network'}
 if($apply.IndexOf('Get-ExcelDailyUpdateBlocker') -gt $apply.IndexOf('Copy-Item -LiteralPath (Join-Path $stage $name)')){throw 'Installer must check Excel before binary replacement'}
 if($apply -notmatch 'ROLLBACK_FAILED' -or $apply -notmatch 'UPDATE_DEFERRED'){throw 'Installer needs rollback and deferral paths'}
 $cfg=Get-Content (Join-Path $PSScriptRoot '..\settings.json') -Raw -Encoding UTF8|ConvertFrom-Json
 if($cfg.auto_update.enabled -ne $true -or [int]$cfg.auto_update.check_interval_hours -ne 6){throw 'Automatic update defaults unexpectedly changed'}
 Write-Output 'UPDATE_SAFETY_TESTS_OK'
}finally{Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue}
